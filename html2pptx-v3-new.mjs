#!/usr/bin/env node
/**
 * html2pptx.mjs — 通用 HTML 幻灯片转 PPTX 工具
 *
 * 核心思路：
 *   1. 用 Puppeteer 打开 HTML，等待 JS 图表渲染完成
 *   2. 在浏览器中遍历 DOM，用 getComputedStyle + getBoundingClientRect 获取精确信息
 *   3. 对 canvas/svg 图表元素截图，保存为图片
 *   4. 用 PptxGenJS 将文本、形状、图片映射到 PPTX
 *
 * 用法: node html2pptx.mjs <input.html> [output.pptx]
 *       node html2pptx.mjs <directory_of_htmls> [output.pptx]
 */

import fs from 'fs';
import path from 'path';
import puppeteer from 'puppeteer';
import PptxGenJS from 'pptxgenjs';

// ─── 常量 ────────────────────────────────────────────────────────────────────
const SLIDE_W_PX = 1280;
const SLIDE_H_PX = 720;
const SLIDE_W_IN = 13.333;
const SLIDE_H_IN = 7.5;
const PX2IN = SLIDE_W_IN / SLIDE_W_PX;

// ─── 浏览器端脚本：提取 DOM 元素信息 ─────────────────────────────────────────

/**
 * 在浏览器上下文中执行。
 * 递归遍历 .slide-container 的 DOM 树，提取所有可见元素信息。
 * 关键改进：
 *   - rgba 颜色通过 canvas 2d 实际混合计算，得到最终显示色
 *   - 文本从父元素获取对齐信息
 *   - 图标截图而非字符替代
 */
function extractElementsScript(containerSelector) {
  const container = document.querySelector(containerSelector);
  if (!container) return { elements: [], containerRect: null };

  const containerRect = container.getBoundingClientRect();
  const elements = [];

  // ── 颜色工具 ──

  /** 解析 rgb/rgba 字符串为 [r, g, b, a] */
  function parseRGBA(color) {
    if (!color || color === 'transparent') return null;
    const m = color.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*(?:,\s*([\d.]+))?\s*\)/);
    if (!m) return null;
    return [parseInt(m[1]), parseInt(m[2]), parseInt(m[3]), m[4] !== undefined ? parseFloat(m[4]) : 1];
  }

  /** RGBA 混合到背景色 */
  function blendColor(fg, bgHex) {
    if (!fg) return null;
    const [r, g, b, a] = fg;
    if (a >= 0.99) return rgbToHex(r, g, b);
    if (a < 0.005) return null;
    // 解析背景 hex
    let br = 0x12, bg2 = 0x12, bb = 0x12;
    if (bgHex && bgHex.length === 6) {
      br = parseInt(bgHex.substring(0, 2), 16);
      bg2 = parseInt(bgHex.substring(2, 4), 16);
      bb = parseInt(bgHex.substring(4, 6), 16);
    }
    const mr = Math.round(r * a + br * (1 - a));
    const mg = Math.round(g * a + bg2 * (1 - a));
    const mb = Math.round(b * a + bb * (1 - a));
    return rgbToHex(mr, mg, mb);
  }

  function rgbToHex(r, g, b) {
    return ((r << 16) | (g << 8) | b).toString(16).padStart(6, '0');
  }

  function colorToHex(color) {
    const rgba = parseRGBA(color);
    if (!rgba) return null;
    return rgbToHex(rgba[0], rgba[1], rgba[2]);
  }

  // ── 元素判断 ──

  function isChartElement(el) {
    const tag = (el.tagName || '').toUpperCase();
    if (tag === 'CANVAS') return true;
    if (tag === 'SVG') {
      // 排除 Font Awesome SVG 图标（通常很小）
      const rect = el.getBoundingClientRect();
      if (rect.width > 50 && rect.height > 50) return true;
      return false;
    }
    return false;
  }

  function isChartContainer(el) {
    if (!el.querySelector) return false;
    const rect = el.getBoundingClientRect();
    // 排除太大的容器（宽度 >= 85% 的全宽容器不应整体截图）
    if (rect.width >= containerRect.width * 0.85) return false;
    // 只检查直接子元素（不用 querySelector 深度搜索，避免匹配到祖先容器）
    let directChart = null;
    for (const child of el.children) {
      const ctag = (child.tagName || '').toUpperCase();
      if (ctag === 'CANVAS') { directChart = child; break; }
      if (ctag === 'SVG' && child.getBoundingClientRect().width > 50) { directChart = child; break; }
    }
    if (!directChart) return false;
    // 检查 canvas/svg 是否占据了容器的大部分面积
    const childRect = directChart.getBoundingClientRect();
    if (childRect.width > rect.width * 0.4 && childRect.height > rect.height * 0.4) return true;
    return false;
  }

  function isFaIcon(el) {
    const cls = (typeof el.className === 'string') ? el.className : '';
    return (el.tagName === 'I' || el.tagName === 'SPAN') &&
           (cls.includes('fa-') || cls.includes('fas ') || cls.includes('far ') || cls.includes('fab '));
  }

  function getDirectText(el) {
    let text = '';
    for (const node of el.childNodes) {
      if (node.nodeType === Node.TEXT_NODE) {
        const t = node.textContent;
        if (t.trim()) text += t;
      } else if (node.nodeType === Node.ELEMENT_NODE && node.tagName === 'BR') {
        text += '\n';
      }
    }
    return text.replace(/[ \t]+/g, ' ').replace(/^\s+|\s+$/g, '').replace(/\n /g, '\n');
  }

  // ── flex/richtext 辅助函数 ──

  /** 内联标签集合 */
  const INLINE_TAGS = new Set(['SPAN', 'STRONG', 'EM', 'B', 'I', 'A', 'MARK', 'SMALL', 'SUB', 'SUP', 'ABBR', 'CODE', 'U', 'S', 'LABEL']);

  /** 检测元素是否为内联元素（span, strong, em, b, i 等） */
  function isInlineElement(el) {
    if (!el || el.nodeType !== Node.ELEMENT_NODE) return false;
    // 先按标签名快速判断
    if (INLINE_TAGS.has(el.tagName)) {
      const d = window.getComputedStyle(el).display;
      return d === 'inline' || d === 'inline-block';
    }
    return false;
  }

  /** 检测元素是否为 flex/grid 容器 */
  function isFlexOrGrid(el) {
    const d = window.getComputedStyle(el).display;
    return d === 'flex' || d === 'inline-flex' || d === 'grid' || d === 'inline-grid';
  }

  /** 检查元素是否包含内联子元素（span, strong, em 等） */
  function hasInlineChildren(el) {
    for (const child of el.children) {
      if (isInlineElement(child) && child.textContent.trim()) return true;
    }
    return false;
  }

  /**
   * 提取元素的富文本（包括内联 span），返回 textParts 数组
   * 例如: <div class="stat-value">5.16 <span>Trillion</span></div>
   * 返回: [{text: "5.16 ", fontSize: 48, ...}, {text: "Trillion", fontSize: 24, ...}]
   */
  function extractRichText(el) {
    const parts = [];
    const baseStyle = window.getComputedStyle(el);
    for (const node of el.childNodes) {
      if (node.nodeType === Node.TEXT_NODE) {
        const t = node.textContent.replace(/\s+/g, ' ');
        if (t.trim()) {
          parts.push({
            text: t,
            color: colorToHex(baseStyle.color),
            fontSize: parseFloat(baseStyle.fontSize),
            fontWeight: baseStyle.fontWeight,
            fontFamily: baseStyle.fontFamily,
          });
        }
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        if (node.tagName === 'BR') {
          parts.push({ text: '\n' });
        } else if (isInlineElement(node)) {
          const inlineStyle = window.getComputedStyle(node);
          const inlineText = node.textContent.replace(/[ \t]+/g, ' ').trim();
          if (inlineText) {
            parts.push({
              text: inlineText,
              color: colorToHex(inlineStyle.color),
              fontSize: parseFloat(inlineStyle.fontSize),
              fontWeight: inlineStyle.fontWeight,
              fontFamily: inlineStyle.fontFamily,
            });
          }
        }
      }
    }
    return parts;
  }

  /**
   * 对于 flex 容器，提取所有子元素的精确位置
   * 使用 Range API 获取文本节点的真实坐标，而不是父容器的坐标
   */
  function extractFlexChildTexts(el) {
    const items = [];
    const elStyle = window.getComputedStyle(el);
    for (const node of el.childNodes) {
      if (node.nodeType === Node.TEXT_NODE) {
        const t = node.textContent.replace(/[ \t]+/g, ' ').trim();
        if (!t) continue;
        const range = document.createRange();
        range.selectNode(node);
        const r = range.getBoundingClientRect();
        if (r.width > 0 && r.height > 0) {
          items.push({
            text: t,
            x: r.left - containerRect.left,
            y: r.top - containerRect.top,
            w: r.width,
            h: r.height,
            color: colorToHex(elStyle.color),
            fontSize: parseFloat(elStyle.fontSize),
            fontWeight: elStyle.fontWeight,
            fontFamily: elStyle.fontFamily,
            fontStyle: elStyle.fontStyle,
            textAlign: 'left',
            lineHeight: parseFloat(elStyle.lineHeight) || parseFloat(elStyle.fontSize) * 1.2,
            letterSpacing: parseFloat(elStyle.letterSpacing) || 0,
            textTransform: elStyle.textTransform,
          });
        }
      } else if (node.nodeType === Node.ELEMENT_NODE && isInlineElement(node)) {
        const inlineText = node.textContent.trim();
        if (!inlineText) continue;
        const inlineRect = node.getBoundingClientRect();
        const inlineNodeStyle = window.getComputedStyle(node);
        if (inlineRect.width > 0 && inlineRect.height > 0) {
          items.push({
            text: inlineText,
            x: inlineRect.left - containerRect.left,
            y: inlineRect.top - containerRect.top,
            w: inlineRect.width,
            h: inlineRect.height,
            color: colorToHex(inlineNodeStyle.color),
            fontSize: parseFloat(inlineNodeStyle.fontSize),
            fontWeight: inlineNodeStyle.fontWeight,
            fontFamily: inlineNodeStyle.fontFamily,
            fontStyle: inlineNodeStyle.fontStyle,
            textAlign: 'left',
            lineHeight: parseFloat(inlineNodeStyle.lineHeight) || parseFloat(inlineNodeStyle.fontSize) * 1.2,
            letterSpacing: parseFloat(inlineNodeStyle.letterSpacing) || 0,
            textTransform: inlineNodeStyle.textTransform,
          });
        }
      }
    }
    return items;
  }

  // ── 递归遍历 ──

  function traverse(el, depth, parentBgHex) {
    if (!el || !el.getBoundingClientRect) return;

    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);

    if (style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0') return;
    if (rect.width <= 0 || rect.height <= 0) return;

    const x = rect.left - containerRect.left;
    const y = rect.top - containerRect.top;
    const w = rect.width;
    const h = rect.height;

    if (x + w < -5 || y + h < -5 || x > containerRect.width + 5 || y > containerRect.height + 5) return;

    const tag = el.tagName;

    // 计算当前元素的实际背景色（混合后）
    const bgRGBA = parseRGBA(style.backgroundColor);
    const bgBlended = blendColor(bgRGBA, parentBgHex);
    const currentBgHex = bgBlended || parentBgHex;

    // ── 图表容器：标记截图，不递归 ──
    if (isChartContainer(el) || isChartElement(el)) {
      elements.push({ type: 'chart', x, y, w, h, tag, id: el.id || '' });
      return;
    }

    // ── Font Awesome 图标：标记截图 ──
    if (isFaIcon(el)) {
      elements.push({
        type: 'icon',
        x, y, w, h,
        color: colorToHex(style.color),
        fontSize: parseFloat(style.fontSize),
      });
      return;
    }

    // ── 背景色 / 边框 → 形状 ──
    const hasBg = bgRGBA && bgRGBA[3] > 0.005 && bgBlended !== parentBgHex;
    const borderTop = parseFloat(style.borderTopWidth) || 0;
    const borderRight = parseFloat(style.borderRightWidth) || 0;
    const borderBottom = parseFloat(style.borderBottomWidth) || 0;
    const borderLeft = parseFloat(style.borderLeftWidth) || 0;
    const hasBorder = borderTop > 0 || borderRight > 0 || borderBottom > 0 || borderLeft > 0;

    if (hasBg || hasBorder) {
      const borderInfo = {};
      if (hasBorder) {
        const sides = [];
        if (borderTop > 0) sides.push({ side: 'top', width: borderTop, color: colorToHex(style.borderTopColor) });
        if (borderRight > 0) sides.push({ side: 'right', width: borderRight, color: colorToHex(style.borderRightColor) });
        if (borderBottom > 0) sides.push({ side: 'bottom', width: borderBottom, color: colorToHex(style.borderBottomColor) });
        if (borderLeft > 0) sides.push({ side: 'left', width: borderLeft, color: colorToHex(style.borderLeftColor) });
        borderInfo.sides = sides;
      }

      elements.push({
        type: 'shape',
        x, y, w, h,
        fill: hasBg ? bgBlended : null,
        border: hasBorder ? borderInfo : null,
      });
    }

    // ── flex/grid 容器中的文本：使用 Range API 获取精确位置 ──
    // 例如: <div class="model-item" style="display:flex; justify-content:space-between">
    //         <span class="model-rank">01</span> MiniMax M2.5 (CN)
    //       </div>
    if (isFlexOrGrid(el)) {
      const hasTextOrSpan = getDirectText(el) || hasInlineChildren(el);
      if (hasTextOrSpan) {
        const flexItems = extractFlexChildTexts(el);
        for (const item of flexItems) {
          elements.push({
            type: 'text',
            x: item.x, y: item.y, w: item.w, h: item.h,
            text: item.text,
            color: item.color,
            fontSize: item.fontSize,
            fontWeight: item.fontWeight,
            fontFamily: item.fontFamily,
            fontStyle: item.fontStyle,
            textAlign: item.textAlign,
            lineHeight: item.lineHeight,
            letterSpacing: item.letterSpacing,
            textTransform: item.textTransform,
            parentDisplay: style.display,
            parentAlignItems: style.alignItems || '',
            parentJustifyContent: style.justifyContent || '',
          });
        }
        // 递归非内联文本的子元素
        for (const child of el.children) {
          if (!isInlineElement(child) || !child.textContent.trim()) {
            traverse(child, depth + 1, currentBgHex);
          }
        }
        return;
      }
    }

    // ── 富文本（非 flex 容器，包含内联 span） ──
    // 例如: <div class="stat-value">5.16 <span>Trillion</span></div>
    // 所有 parts 放在同一个文本框内，用 PptxGenJS 多样式文本 API
    if (hasInlineChildren(el)) {
      const richParts = extractRichText(el);
      if (richParts.length > 0 && richParts.some(p => p.text && p.text.trim())) {
        const parent = el.parentElement;
        const parentStyle = parent ? window.getComputedStyle(parent) : {};
        const parentRect = parent ? parent.getBoundingClientRect() : rect;
        // 扩展宽度到父容器剩余宽度，避免文字被截断
        const richW = Math.max(w, parentRect.width - (rect.left - parentRect.left) - 4);
        const lineH = parseFloat(style.lineHeight) || parseFloat(style.fontSize) * 1.2;
        const fSize = parseFloat(style.fontSize);
        const brCount = richParts.filter(p => p.text === '\n').length;

        // 高度估算：基于文本内容估算所需行数，给 PPTX 1.4x 安全边际
        const totalChars = richParts.reduce((sum, p) => sum + (p.text === '\n' ? 0 : (p.text || '').length), 0);
        const avgCharW = fSize * 0.75;
        const charsPerLine = Math.max(1, Math.floor(richW / avgCharW));
        const textLines = Math.ceil(totalChars / charsPerLine) + brCount;
        const estimatedH = Math.max(lineH * 2, textLines * lineH * 1.4);
        // 不超过父容器剩余高度
        const parentRemainH = parentRect.bottom - rect.top;
        const richH = Math.max(h, Math.min(estimatedH, parentRemainH));

        elements.push({
          type: 'richtext', x, y, w: richW, h: richH,
          brCount,
          parts: richParts,
          color: colorToHex(style.color),
          fontSize: parseFloat(style.fontSize),
          fontWeight: style.fontWeight,
          fontFamily: style.fontFamily,
          fontStyle: style.fontStyle,
          textAlign: style.textAlign,
          lineHeight: parseFloat(style.lineHeight) || parseFloat(style.fontSize) * 1.2,
          letterSpacing: parseFloat(style.letterSpacing) || 0,
          textTransform: style.textTransform,
          parentDisplay: parentStyle.display || '',
          parentAlignItems: parentStyle.alignItems || '',
          parentJustifyContent: parentStyle.justifyContent || '',
        });

        // 递归非内联元素的子元素
        for (const child of el.children) {
          if (!isInlineElement(child)) {
            traverse(child, depth + 1, currentBgHex);
          }
        }
        return;
      }
    }

    // ── 普通直接文本 ──
    const directText = getDirectText(el);
    if (directText) {
      const parent = el.parentElement;
      const parentStyle = parent ? window.getComputedStyle(parent) : {};
      const parentRect = parent ? parent.getBoundingClientRect() : rect;

      const textW = Math.max(w, parentRect.width - (rect.left - parentRect.left) - 4);
      const fsize = parseFloat(style.fontSize);
      const parentRemainH = parentRect.bottom - rect.top;
      const textH = fsize >= 24 ? Math.max(h, parentRemainH) : h;

      elements.push({
        type: 'text',
        x, y, w: textW, h: textH,
        text: directText,
        color: colorToHex(style.color),
        fontSize: parseFloat(style.fontSize),
        fontWeight: style.fontWeight,
        fontFamily: style.fontFamily,
        fontStyle: style.fontStyle,
        textAlign: style.textAlign,
        lineHeight: parseFloat(style.lineHeight) || parseFloat(style.fontSize) * 1.2,
        letterSpacing: parseFloat(style.letterSpacing) || 0,
        textTransform: style.textTransform,
        parentDisplay: parentStyle.display || '',
        parentAlignItems: parentStyle.alignItems || '',
        parentJustifyContent: parentStyle.justifyContent || '',
      });
    }

    // ── 图片 ──
    if (tag === 'IMG') {
      elements.push({ type: 'image', x, y, w, h, src: el.src });
      return;
    }

    // ── 递归子元素 ──
    for (const child of el.children) {
      traverse(child, depth + 1, currentBgHex);
    }
  }

  const containerBg = colorToHex(window.getComputedStyle(container).backgroundColor) || '121212';
  traverse(container, 0, containerBg);

  return {
    elements,
    containerRect: {
      width: containerRect.width,
      height: containerRect.height,
      bgColor: containerBg,
    },
  };
}

// ─── 截图图表和图标元素 ──────────────────────────────────────────────────────

async function screenshotElements(page, elements, containerSelector, outputDir, type) {
  const results = [];
  const containerRect = await page.$eval(containerSelector, el => {
    const r = el.getBoundingClientRect();
    return { left: r.left, top: r.top };
  });

  for (let i = 0; i < elements.length; i++) {
    const el = elements[i];
    const filename = `${type}_${i}.png`;
    const filepath = path.join(outputDir, filename);

    const clipX = Math.max(0, containerRect.left + el.x);
    const clipY = Math.max(0, containerRect.top + el.y);

    try {
      await page.screenshot({
        path: filepath,
        clip: { x: clipX, y: clipY, width: Math.max(1, el.w), height: Math.max(1, el.h) },
      });
      results.push({ ...el, imagePath: filepath });
      console.log(`  📸 截图 ${type} #${i}: ${Math.round(el.w)}x${Math.round(el.h)}px`);
    } catch (e) {
      console.warn(`  ⚠️ 截图失败 ${type} #${i}: ${e.message}`);
    }
  }
  return results;
}

// ─── PptxGenJS 渲染 ──────────────────────────────────────────────────────────

function px2in(px) { return px * PX2IN; }
function pt2in(pt) { return pt / 72; }

function buildPptxSlide(pres, elementsData, chartImages, iconImages) {
  const slide = pres.addSlide();
  const { elements, containerRect } = elementsData;

  if (containerRect.bgColor) {
    slide.background = { color: containerRect.bgColor.toUpperCase() };
  }

  // 构建截图查找 map（用坐标做 key）
  const imgMap = new Map();
  for (const ci of [...chartImages, ...iconImages]) {
    imgMap.set(`${Math.round(ci.x)},${Math.round(ci.y)}`, ci);
  }

  for (const el of elements) {
    const x = px2in(el.x);
    const y = px2in(el.y);
    const w = px2in(el.w);
    const h = px2in(el.h);

    if (w < 0.01 || h < 0.01) continue;
    if (x + w < 0 || y + h < 0 || x > SLIDE_W_IN || y > SLIDE_H_IN) continue;

    switch (el.type) {
      case 'shape':    renderShape(slide, pres, el, x, y, w, h); break;
      case 'text':     renderText(slide, el, x, y, w, h); break;
      case 'richtext': renderRichText(slide, el, x, y, w, h); break;
      case 'chart':
      case 'icon':     renderScreenshot(slide, el, x, y, w, h, imgMap); break;
      case 'image':    renderImage(slide, el, x, y, w, h); break;
    }
  }
  return slide;
}

function renderShape(slide, pres, el, x, y, w, h) {
  const opts = { x, y, w, h };

  if (el.fill) {
    opts.fill = { color: el.fill.toUpperCase() };
  }

  if (el.border && el.border.sides) {
    const sides = el.border.sides;

    if (sides.length === 4) {
      const s = sides[0];
      opts.line = { color: (s.color || '333333').toUpperCase(), width: Math.max(s.width * 0.75, 0.5) };
      slide.addShape(pres.ShapeType.rect, opts);
    } else {
      // 先画填充
      if (el.fill) {
        slide.addShape(pres.ShapeType.rect, { x, y, w, h, fill: { color: el.fill.toUpperCase() } });
      }
      // 单独画每条边
      for (const s of sides) {
        const lineColor = (s.color || '333333').toUpperCase();
        if (s.width >= 2.5) {
          // 粗边框用矩形
          const bw = px2in(s.width);
          const rects = {
            left:   { x, y, w: bw, h },
            right:  { x: x + w - bw, y, w: bw, h },
            top:    { x, y, w, h: bw },
            bottom: { x, y: y + h - bw, w, h: bw },
          };
          if (rects[s.side]) {
            slide.addShape(pres.ShapeType.rect, { ...rects[s.side], fill: { color: lineColor } });
          }
        } else {
          // 细边框用线条
          const lines = {
            top:    { x, y, w, h: 0 },
            bottom: { x, y: y + h, w, h: 0 },
            left:   { x, y, w: 0, h },
            right:  { x: x + w, y, w: 0, h },
          };
          if (lines[s.side]) {
            slide.addShape(pres.ShapeType.line, {
              ...lines[s.side],
              line: { color: lineColor, width: Math.max(s.width * 0.75, 0.5) },
            });
          }
        }
      }
    }
  } else {
    slide.addShape(pres.ShapeType.rect, opts);
  }
}

function renderText(slide, el, x, y, w, h) {
  let text = el.text;
  if (el.textTransform === 'uppercase') text = text.toUpperCase();
  else if (el.textTransform === 'lowercase') text = text.toLowerCase();

  const fontSize = el.fontSize ? el.fontSize * 0.75 : 12;
  const color = (el.color || 'ffffff').toUpperCase();
  const bold = el.fontWeight && (parseInt(el.fontWeight) >= 700 || el.fontWeight === 'bold');
  const italic = el.fontStyle === 'italic';

  // 推断字体
  let fontFace = 'Arial';
  const ff = (el.fontFamily || '').toLowerCase();
  if (ff.includes('mono') || ff.includes('consolas') || ff.includes('jetbrains') || ff.includes('courier')) {
    fontFace = 'Consolas';
  } else if (ff.includes('yahei')) {
    fontFace = 'Microsoft YaHei';
  }

  // 对齐
  let align = 'left';
  if (el.textAlign === 'center') align = 'center';
  else if (el.textAlign === 'right' || el.textAlign === 'end') align = 'right';
  else if (el.textAlign === '-webkit-match-parent' || el.textAlign === 'start') align = 'left';

  let valign = 'top';
  if (el.parentDisplay?.includes('flex') && el.parentAlignItems === 'center') valign = 'middle';

  // 对大字号文本增加高度安全边际
  // getBoundingClientRect 返回的是文字实际渲染高度，但 PPTX 中文本框需要额外的行高和 padding 空间
  let adjustedH = h;
  if (fontSize >= 14) {
    // 确保至少有 fontSize * 2.5 的高度（pt 转 inches），给单行文本足够空间
    const minH = pt2in(fontSize * 2.5);
    adjustedH = Math.max(h, minH);
  }

  const opts = {
    x, y, w, h: adjustedH,
    fontSize, fontFace, color, bold, italic,
    align, valign,
    wrap: true,
    lang: 'zh-CN',
    margin: 0, // 去掉默认 margin，精确控制位置
  };

  // 对大字号文本始终启用 shrink，防止不同系统字体渲染差异导致溢出
  if (fontSize >= 16) {
    opts.fit = 'shrink';
  } else {
    // 小字号文本也检查溢出
    const avgCharWidth = fontSize * 0.55;
    const lineWidthPt = w * 72;
    const charsPerLine = Math.floor(lineWidthPt / avgCharWidth) || 1;
    const textLines = Math.ceil(text.length / charsPerLine);
    const lineHeightPt = fontSize * 1.4;
    const availableLines = Math.floor((adjustedH * 72) / lineHeightPt) || 1;
    if (textLines > availableLines) {
      opts.fit = 'shrink';
    }
  }

  // 行间距
  if (el.lineHeight && el.fontSize) {
    const ratio = el.lineHeight / el.fontSize;
    if (ratio > 0.8 && ratio < 3) {
      opts.lineSpacingMultiple = ratio;
    }
  }

  if (el.letterSpacing > 0) {
    opts.charSpacing = el.letterSpacing * 0.75;
  }

  slide.addText(text, opts);
}

function renderRichText(slide, el, x, y, w, h) {
  const fontSize = el.fontSize ? el.fontSize * 0.75 : 12;
  const color = (el.color || 'ffffff').toUpperCase();
  const bold = el.fontWeight && (parseInt(el.fontWeight) >= 700 || el.fontWeight === 'bold');
  const italic = el.fontStyle === 'italic';

  let fontFace = 'Arial';
  const ff = (el.fontFamily || '').toLowerCase();
  if (ff.includes('mono') || ff.includes('consolas') || ff.includes('jetbrains') || ff.includes('courier')) {
    fontFace = 'Consolas';
  } else if (ff.includes('yahei')) {
    fontFace = 'Microsoft YaHei';
  }

  let align = 'left';
  if (el.textAlign === 'center') align = 'center';
  else if (el.textAlign === 'right' || el.textAlign === 'end') align = 'right';

  let valign = 'top';
  if (el.parentDisplay?.includes('flex') && el.parentAlignItems === 'center') valign = 'middle';

  // 大字号增加高度安全边际
  let adjustedH = h;
  if (fontSize >= 14) {
    const minH = pt2in(fontSize * 2.5);
    adjustedH = Math.max(h, minH);
  }

  // 构建多样式文本 parts
  const textParts = el.parts.map(part => {
    if (part.text === '\n') {
      return { text: '\n', options: { fontSize, fontFace, breakType: 'break' } };
    }
    let partText = part.text;
    if (el.textTransform === 'uppercase') partText = partText.toUpperCase();
    else if (el.textTransform === 'lowercase') partText = partText.toLowerCase();

    const partFontSize = part.fontSize ? part.fontSize * 0.75 : fontSize;
    const partColor = (part.color || color).toUpperCase();
    const partBold = part.fontWeight && (parseInt(part.fontWeight) >= 700 || part.fontWeight === 'bold');

    let partFontFace = fontFace;
    const pff = (part.fontFamily || '').toLowerCase();
    if (pff.includes('mono') || pff.includes('consolas') || pff.includes('jetbrains') || pff.includes('courier')) {
      partFontFace = 'Consolas';
    }

    return {
      text: partText,
      options: {
        fontSize: partFontSize,
        fontFace: partFontFace,
        color: partColor,
        bold: partBold || bold,
        italic,
        lang: 'zh-CN',
      },
    };
  });

  // 先计算 lineSpacingMultiple，shrink 判断需要用到
  let lineSpacingMultiple = 1.2;
  if (el.lineHeight && el.fontSize) {
    const ratio = el.lineHeight / el.fontSize;
    if (ratio > 0.8 && ratio < 3) lineSpacingMultiple = ratio;
  }

  const opts = {
    x, y, w, h: adjustedH,
    align, valign,
    wrap: true,
    margin: 0,
    lang: 'zh-CN',
    lineSpacingMultiple,
  };

  // shrink 策略：
  // - 大字号（≥ 16pt）：始终启用 shrink，防止单行溢出
  // - 小字号：仅当估算内容超出文本框时才启用 shrink，否则用 wrap 自然换行
  if (fontSize >= 16) {
    opts.fit = 'shrink';
  } else {
    // 估算内容是否超出文本框
    const totalText = el.parts.reduce((s, p) => s + (p.text || ''), '');
    const brCount = el.parts.filter(p => p.text === '\n').length;
    const wInPt = w * 72;  // inches to points
    const hInPt = adjustedH * 72;
    // 中英文混排：估算平均字符宽度，中文约 1.0x fontSize，英文约 0.55x
    const cjkCount = (totalText.match(/[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]/g) || []).length;
    const nonCjkCount = totalText.length - cjkCount - brCount;
    const textWidthPt = cjkCount * fontSize * 1.0 + nonCjkCount * fontSize * 0.55;
    const estLines = Math.ceil(textWidthPt / wInPt) + brCount;
    const neededH = estLines * fontSize * lineSpacingMultiple;
    if (neededH > hInPt * 0.9) {
      opts.fit = 'shrink';
    }
  }

  slide.addText(textParts, opts);
}

function renderScreenshot(slide, el, x, y, w, h, imgMap) {
  const key = `${Math.round(el.x)},${Math.round(el.y)}`;
  const info = imgMap.get(key);
  if (info && info.imagePath && fs.existsSync(info.imagePath)) {
    slide.addImage({ path: info.imagePath, x, y, w, h });
  }
}

function renderImage(slide, el, x, y, w, h) {
  if (!el.src) return;
  try {
    if (el.src.startsWith('data:')) {
      slide.addImage({ data: el.src, x, y, w, h });
    } else {
      slide.addImage({ path: el.src, x, y, w, h });
    }
  } catch (e) {
    console.warn(`  ⚠️ 无法添加图片: ${el.src.substring(0, 60)}...`);
  }
}

// ─── 主流程 ───────────────────────────────────────────────────────────────────

async function convertHtmlToPptx(htmlFiles, outputPath) {
  console.log(`🚀 启动浏览器...`);
  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-web-security'],
  });

  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'WIDE', width: SLIDE_W_IN, height: SLIDE_H_IN });
  pres.layout = 'WIDE';

  const tmpDir = path.join(path.dirname(outputPath), '.html2pptx_tmp');
  if (!fs.existsSync(tmpDir)) fs.mkdirSync(tmpDir, { recursive: true });

  for (let i = 0; i < htmlFiles.length; i++) {
    const htmlFile = htmlFiles[i];
    console.log(`\n📄 处理第 ${i + 1}/${htmlFiles.length} 页: ${path.basename(htmlFile)}`);

    const page = await browser.newPage();
    await page.setViewport({ width: 1920, height: 1080, deviceScaleFactor: 2 });

    const fileUrl = 'file://' + path.resolve(htmlFile);
    await page.goto(fileUrl, { waitUntil: 'networkidle0', timeout: 30000 }).catch(() => {
      console.log('  ⚠️ 部分网络资源加载超时，继续处理...');
    });

    // 等待图表渲染
    await new Promise(r => setTimeout(r, 2000));

    // 提取元素信息
    console.log('  🔍 提取 DOM 元素信息...');
    const containerSelector = '.slide-container';
    const elementsData = await page.evaluate(extractElementsScript, containerSelector);

    if (!elementsData.containerRect) {
      console.log('  ⚠️ 未找到 .slide-container，跳过此页');
      await page.close();
      continue;
    }

    console.log(`  📊 发现 ${elementsData.elements.length} 个元素`);
    const typeCounts = {};
    for (const el of elementsData.elements) {
      typeCounts[el.type] = (typeCounts[el.type] || 0) + 1;
    }
    console.log(`  📋 类型分布: ${JSON.stringify(typeCounts)}`);

    // 截图图表元素
    const charts = elementsData.elements.filter(e => e.type === 'chart');
    let chartImages = [];
    if (charts.length > 0) {
      console.log(`  📸 截图 ${charts.length} 个图表元素...`);
      chartImages = await screenshotElements(page, charts, containerSelector, tmpDir, `s${i}_chart`);
    }

    // 截图图标元素
    const icons = elementsData.elements.filter(e => e.type === 'icon');
    let iconImages = [];
    if (icons.length > 0) {
      console.log(`  📸 截图 ${icons.length} 个图标元素...`);
      iconImages = await screenshotElements(page, icons, containerSelector, tmpDir, `s${i}_icon`);
    }

    // 构建 PPTX 幻灯片
    console.log('  🎨 生成 PPTX 幻灯片...');
    buildPptxSlide(pres, elementsData, chartImages, iconImages);

    await page.close();
  }

  await browser.close();

  console.log(`\n💾 保存 PPTX...`);
  await pres.writeFile({ fileName: outputPath });
  console.log(`✅ 转换完成！输出文件: ${outputPath}`);
  console.log(`   幻灯片数量: ${htmlFiles.length}`);

  // 清理临时文件
  try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch (e) {}
}

// ─── CLI ──────────────────────────────────────────────────────────────────────

async function main() {
  const args = process.argv.slice(2);
  if (args.length < 1) {
    console.log(`
html2pptx — 通用 HTML 幻灯片转 PPTX 工具

用法:
  node html2pptx.mjs <input.html> [output.pptx]
  node html2pptx.mjs <directory> [output.pptx]

说明:
  - 单个 HTML 文件: 转换为单页 PPTX
  - 目录: 将目录下所有 .html 文件按文件名排序，合并为多页 PPTX
  - Chart.js / D3.js 图表自动截图并以图片形式嵌入
  - Font Awesome 图标自动截图嵌入
  - 要求 HTML 包含 .slide-container 容器（1280x720px）
`);
    process.exit(1);
  }

  const inputPath = args[0];
  let outputPath = args[1];
  let htmlFiles = [];

  const stat = fs.statSync(inputPath);
  if (stat.isDirectory()) {
    const files = fs.readdirSync(inputPath)
      .filter(f => f.endsWith('.html') || f.endsWith('.htm'))
      .sort();
    htmlFiles = files.map(f => path.join(inputPath, f));
    if (!outputPath) outputPath = path.join(inputPath, 'output.pptx');
  } else {
    htmlFiles = [inputPath];
    if (!outputPath) outputPath = inputPath.replace(/\.html?$/i, '') + '.pptx';
  }

  if (htmlFiles.length === 0) {
    console.error('未找到 HTML 文件');
    process.exit(1);
  }

  console.log(`📁 待转换文件: ${htmlFiles.length} 个`);
  await convertHtmlToPptx(htmlFiles, outputPath);
}

main().catch(err => {
  console.error('❌ 转换失败:', err);
  process.exit(1);
});
