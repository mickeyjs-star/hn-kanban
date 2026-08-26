/*
 * docx_builder.js — 零售经理期间对比分析 .docx 文档构造（浏览器 / Node 通用）
 * 依赖全局 Docx（docx 库命名空间：window.Docx 或 global.Docx）。
 * buildManagerDocument(model) 返回 docx Document 对象，由调用方用 Packer 导出。
 *
 * model = {
 *   manager, curL, baseL,          // 经理名、本期/对比期标签
 *   tCur, tBase, tDiff, tPct,     // 整体潜客：本期、对比期、变化量、变化%
 *   chan: [{name,b,c,dd,p}],      // 各渠道（b=对比期 c=本期 dd=c-b p=环比%或null=新增）
 *   rows: [{name,city,b,c,dd}],   // 各门店（已排序）
 *   up, down, flat,               // 上升/下降/持平门店数
 *   reasonRows: [{name,city,b,c,dd,upch:[{name,cb,cc,cdd}],downch:[{name,cb,cc,cdd}]}],
 *   note                          // 可选：追加说明段落（字符串）
 * }
 */
(function (root) {
  function buildManagerDocument(model) {
    const Docx = (typeof window !== 'undefined' && window.Docx) ? window.Docx : (typeof global !== 'undefined' ? global.Docx : null);
    const {
      Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
      AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType, ShadingType
    } = Docx;

    const FONT = 'Microsoft YaHei';
    const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
    const borders = { top: border, bottom: border, left: border, right: border };
    const hdrFill = { fill: 'D5E8F0', type: ShadingType.CLEAR };
    const alignR = AlignmentType.RIGHT;

    function cell(text, w, opts) {
      opts = opts || {};
      return new TableCell({
        borders, width: { size: w, type: WidthType.DXA },
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        shading: opts.fill ? { ...hdrFill, fill: opts.fill } : undefined,
        children: [new Paragraph({
          alignment: opts.right ? alignR : AlignmentType.LEFT,
          children: [new TextRun({ text: String(text), bold: !!opts.bold, font: FONT, size: opts.size || 20 })]
        })]
      });
    }
    function table(headers, rowsData, colW, rightCols) {
      rightCols = rightCols || [];
      const header = new TableRow({ children: headers.map((h, i) => cell(h, colW[i], { bold: true, fill: 'D5E8F0', right: rightCols.includes(i) })) });
      const body = rowsData.map(r => new TableRow({ children: r.map((v, i) => cell(v, colW[i], { right: rightCols.includes(i) })) }));
      return new Table({ width: { size: colW.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: colW, rows: [header, ...body] });
    }

    const pctStr = p => p === null ? '新增' : (p === 0 ? '—' : (p >= 0 ? '+' : '') + p.toFixed(1) + '%');
    const mark = dd => dd > 0 ? '▲上升' : (dd < 0 ? '▼下降' : '＝持平');

    const { manager, curL, baseL, tCur, tBase, tDiff, tPct, chan, rows, up, down, flat, reasonRows, note } = model;

    const children = [];
    // 标题
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: `${manager} · ${curL} vs ${baseL} 潜客变化分析`, font: FONT, bold: true })] }));
    children.push(new Paragraph({ children: [new TextRun({ text: `口径：数据每日表为“当月累计”，对比期为本月所选区间的上一月同期（同日序号对齐）；两区间快照相减得区间合计，全部车系合并，仅筛选零售经理 ${manager}（共 ${rows.length} 家门店）。`, font: FONT, size: 18, italics: true })] }));

    // 一、整体
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: '一、整体潜客变化', font: FONT, bold: true })] }));
    children.push(new Paragraph({ children: [new TextRun({ text: `总潜客：${tBase} → ${tCur}（Δ${tDiff >= 0 ? '+' : ''}${tDiff}，${tPct >= 0 ? '+' : ''}${tPct.toFixed(1)}%）`, font: FONT, size: 22, bold: true })] }));
    {
      const sorted = [...chan].sort((a, b) => b.dd - a.dd);
      const headers = ['渠道', `${baseL}`, `${curL}`, '变化', '环比'];
      const data = sorted.map(c => [c.name, c.b, c.c, (c.dd >= 0 ? '+' : '') + c.dd, pctStr(c.p)]);
      const colW = [2400, 1657, 1657, 1657, 1655];
      children.push(table(headers, data, colW, [1, 2, 3, 4]));
    }

    // 二、分门店
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: '二、分门店潜客变化', font: FONT, bold: true })] }));
    children.push(new Paragraph({ children: [new TextRun({ text: `上升门店 ${up} 家，下降门店 ${down} 家，持平 ${flat} 家（按 ${curL} 潜客降序）。`, font: FONT, size: 20 })] }));
    {
      const headers = ['门店', '城市', `${baseL}`, `${curL}`, '变化', '趋势'];
      const data = rows.map(r => [r.name, r.city, r.b, r.c, (r.dd >= 0 ? '+' : '') + r.dd, mark(r.dd)]);
      const colW = [2900, 1100, 1460, 1460, 1460, 646];
      children.push(table(headers, data, colW, [2, 3, 4, 5]));
    }

    // 三、门店渠道拆解与下降原因分析
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: '三、门店渠道拆解与下降原因分析', font: FONT, bold: true })] }));
    const bullet = { config: [{ reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 600, hanging: 300 } } } }] }] };
    reasonRows.forEach(d => {
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun({ text: `${d.name}（${d.city}）— 总潜客 ${d.b} → ${d.c}（${(d.dd >= 0 ? '+' : '') + d.dd}）`, font: FONT, bold: true })] }));
      if (d.upch.length) {
        children.push(new Paragraph({ children: [new TextRun({ text: `上升渠道（${d.upch.length}）：`, font: FONT, size: 20, bold: true })] }));
        d.upch.forEach(c => children.push(new Paragraph({ numbering: { reference: 'bul', level: 0 }, children: [new TextRun({ text: `${c.name} ${c.cb}→${c.cc}（+${c.cdd}）`, font: FONT, size: 20 })] })));
      }
      const rHeaders = ['下降渠道（本月 → 上月，变化量）', '原因 / 改善动作（请填写）'];
      const rData = d.downch.map(c => [`${c.name}（${c.cb}→${c.cc}，${c.cdd > 0 ? '+' : ''}${c.cdd}）`, '（请填写）']);
      children.push(new Paragraph({ children: [new TextRun({ text: `下降原因分析（${d.downch.length} 个下滑渠道）：`, font: FONT, size: 20, bold: true })] }));
      children.push(table(rHeaders, rData, [4500, 4526], []));
    });

    if (note) {
      children.push(new Paragraph({ children: [new TextRun({ text: note, font: FONT, size: 18, italics: true, color: 'B45309' })] }));
    }

    const doc = new Document({
      numbering: bullet,
      styles: {
        default: { document: { run: { font: FONT, size: 20 } } },
        paragraphStyles: [
          { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 32, bold: true, font: FONT }, paragraph: { spacing: { before: 240, after: 160 } } },
          { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 26, bold: true, font: FONT }, paragraph: { spacing: { before: 200, after: 120 } } },
          { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 22, bold: true, font: FONT }, paragraph: { spacing: { before: 120, after: 80 } } }
        ]
      },
      sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } }, children }]
    });
    return doc;
  }

  if (typeof module !== 'undefined' && module.exports) module.exports = { buildManagerDocument };
  if (typeof window !== 'undefined') window.buildManagerDocument = buildManagerDocument;
})(typeof globalThis !== 'undefined' ? globalThis : this);
