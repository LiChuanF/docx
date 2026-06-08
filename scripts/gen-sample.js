/* eslint-disable */
/**
 * 生成"带红头(武文头)+ 命名样式正文"的 docx
 *
 * 红头公文版式(GB/T 9704-2012):
 *   1. 红色大字发文机关(内容可变) — 命名样式 "Govt Red Header"
 *   2. 粗红反线(上)        ─ 段落下边框 24 半磅
 *   3. 细红线(下,紧贴)     ─ 段落下边框 6 半磅
 *   4. 正文...
 *   5. 结尾粗红线           ─ 命名样式 "Govt Closing Rule" 的空段落
 *
 * 关键:红线在 docx 里画给 Word 打开的用户看;mammoth 解析时一定会丢段落边框,
 * 所以前端用 styleMap + 正则,在红头后/结尾位置补对应的 <hr>,CSS 负责画双线和结尾线。
 * 两路独立,不冲突。
 *
 * 运行: node scripts/gen-sample.js  →  public/sample.docx
 */
const fs = require('fs')
const path = require('path')
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  BorderStyle
} = require('docx')

const RED = 'C00000'

const doc = new Document({
  styles: {
    paragraphStyles: [
      // —— 红头大字 ——
      // mammoth 解析时只关心 name,Word 打开也能看到正确样式
      { id: 'GovtRedHeader', name: 'Govt Red Header',
        basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: '方正小标宋简体', size: 108, bold: true, color: RED }, // 108 半磅 = 54pt
        paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 240, after: 120 } }
      },

      // —— 文末结尾粗红线(给前端识别用) ——
      // 这个段落本身在 Word 里靠下边框画线;mammoth 丢边框但保留样式名,
      // 前端 styleMap 把它映射成 <hr class="govt-closing">
      { id: 'GovtClosingRule', name: 'Govt Closing Rule',
        basedOn: 'Normal', next: 'Normal',
        paragraph: {
          spacing: { before: 240 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: RED, space: 1 } }
        }
      },

      // —— 标题级别 ——
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: '方正小标宋简体', size: 44, bold: true },
        paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 240, after: 240 } } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: '黑体', size: 32 },
        paragraph: { spacing: { before: 180, after: 80 } } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: '楷体', size: 32 },
        paragraph: { spacing: { before: 100, after: 40 }, indent: { firstLine: 480 } } }
    ]
  },
  sections: [{
    properties: {
      page: { margin: { top: 2098, right: 1474, bottom: 1985, left: 1587 } }
    },
    children: [
      // ===== 红头大字(机关名,内容可变) =====
      new Paragraph({
        style: 'GovtRedHeader',
        children: [new TextRun({ text: '中华人民共和国应急管理部', characterSpacing: 280 })]
      }),

      // ===== 双红线(给 Word 看;mammoth 会丢,前端 CSS 另外画) =====
      new Paragraph({
        children: [new TextRun({ text: '' })],
        border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: RED, space: 1 } }
      }),
      new Paragraph({
        children: [new TextRun({ text: '' })],
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: RED, space: 1 } },
        spacing: { after: 480 }
      }),

      // ===== 正文标题 =====
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: '金华市金水区地震应急预案' })]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: '1 总则' })]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun({ text: '1.1 编制目的' })]
      }),
      new Paragraph({
        children: [new TextRun({
          text: '贯彻落实习近平总书记关于防范化解重大安全风险和防灾减灾救灾重要论述,依法科学统一、有力有序有效实施地震应急工作,最大程度减少人员伤亡和经济损失,维护社会正常秩序。'
        })],
        indent: { firstLine: 480 }
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun({ text: '1.2 编制依据' })]
      }),
      new Paragraph({
        children: [new TextRun({
          text: '依据《中华人民共和国防震减灾法》《中华人民共和国突发事件应对法》《破坏性地震应急条例》《浙江省防震减灾条例》《浙江省地震应急预案》等。'
        })],
        indent: { firstLine: 480 }
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun({ text: '1.3 适用范围' })]
      }),
      new Paragraph({
        children: [new TextRun({
          text: '本预案适用于本区行政区域内发生地震灾害和区外发生对我区造成重大影响的地震灾害的应对工作及其他相关地震事件的应急处置。'
        })],
        indent: { firstLine: 480 }
      }),

      // ===== 结尾红线 =====
      // 注意:mammoth 会丢弃空段落,所以塞一个不间断空格   占位,
      // 这样段落连同 style-name='Govt Closing Rule' 才会被解析出来
      new Paragraph({
        style: 'GovtClosingRule',
        children: [new TextRun({ text: ' ' })]
      })
    ]
  }]
})

Packer.toBuffer(doc).then(buf => {
  const out = path.join(__dirname, '..', 'public', 'sample.docx')
  fs.writeFileSync(out, buf)
  console.log('生成成功:', out)
  console.log('文件大小:', buf.length, 'bytes')
})
