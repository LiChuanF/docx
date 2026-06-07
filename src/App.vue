<template>
  <div class="app">

    <!-- 顶部操作栏 -->
    <div class="toolbar-top">
      <button @click="loadTemplate">加载红头模板</button>
      <button @click="aiGenerate">AI 生成内容</button>
      <button @click="exportWord">导出 Word</button>
      <button @click="exportMarkdown">导出 MD</button>
    </div>

    <!-- 样式工具栏 -->
    <div v-if="editor" class="toolbar-format">
      <select :value="currentFontFamily" @change="setFontFamily($event.target.value)">
        <option value="">字体</option>
        <option value="仿宋_GB2312, FangSong, serif">仿宋_GB2312(公文)</option>
        <option value="SimSun, serif">宋体</option>
        <option value="SimHei, sans-serif">黑体</option>
        <option value="KaiTi, serif">楷体</option>
      </select>

      <select :value="currentFontSize" @change="setFontSize($event.target.value)">
        <option value="">字号</option>
        <option value="56px">初号 (42pt)</option>
        <option value="48px">小初 (36pt)</option>
        <option value="34px">一号 (26pt)</option>
        <option value="32px">小一 (24pt)</option>
        <option value="29px">二号 (22pt 公文标题)</option>
        <option value="24px">小二 (18pt)</option>
        <option value="21px">三号 (16pt 公文正文)</option>
        <option value="20px">小三 (15pt)</option>
        <option value="19px">四号 (14pt)</option>
        <option value="16px">小四 (12pt)</option>
      </select>

      <input type="color" :value="currentColor" @input="setColor($event.target.value)">

      <button :class="{ active: editor.isActive('bold') }" @click="editor.chain().focus().toggleBold().run()"><b>B</b></button>
      <button :class="{ active: editor.isActive('italic') }" @click="editor.chain().focus().toggleItalic().run()"><i>I</i></button>
      <button :class="{ active: editor.isActive('underline') }" @click="editor.chain().focus().toggleUnderline().run()"><u>U</u></button>

      <button :class="{ active: editor.isActive({ textAlign: 'left' }) }" @click="editor.chain().focus().setTextAlign('left').run()">左</button>
      <button :class="{ active: editor.isActive({ textAlign: 'center' }) }" @click="editor.chain().focus().setTextAlign('center').run()">中</button>
      <button :class="{ active: editor.isActive({ textAlign: 'right' }) }" @click="editor.chain().focus().setTextAlign('right').run()">右</button>

      <span class="divider" />

      <!-- 选中文字才显示的 AI 润色 -->
      <button
        v-if="hasSelection"
        class="ai-polish"
        :disabled="polishing"
        @click="polishSelection"
      >
        {{ polishing ? '润色中…' : '✨ AI 润色选中' }}
      </button>
    </div>

    <!-- 编辑器画布(A4 纸张样式) -->
    <div class="paper">
      <editor-content :editor="editor" />
    </div>
  </div>
</template>

<script>
import { Editor, EditorContent } from '@tiptap/vue-2'
import { Extension } from '@tiptap/core'
import StarterKit from '@tiptap/starter-kit'
import TextAlign from '@tiptap/extension-text-align'
// Tiptap v3 改成命名导出 — default 导出已被移除,误用会拿到 undefined
import { TextStyle, Color, FontFamily, FontSize } from '@tiptap/extension-text-style'
import { Underline } from '@tiptap/extension-underline'

import { saveAs } from 'file-saver'
import * as htmlDocx from 'html-docx-js-typescript'

/* LetterSpacing — v3 内置没有,手写一个挂到 textStyle 上,红头标题用 */
const LetterSpacing = Extension.create({
  name: 'letterSpacing',
  addGlobalAttributes () {
    return [{
      types: ['textStyle'],
      attributes: {
        letterSpacing: {
          default: null,
          parseHTML: el => el.style.letterSpacing || null,
          renderHTML: attrs => attrs.letterSpacing ? { style: `letter-spacing: ${attrs.letterSpacing}` } : {}
        }
      }
    }]
  }
})

export default {
  components: { EditorContent },

  data () {
    return {
      editor: null,
      hasSelection: false,
      polishing: false,
      currentFontFamily: '',
      currentFontSize: '',
      currentColor: '#000000'
    }
  },

  mounted () {
    this.editor = new Editor({
      extensions: [
        StarterKit,
        Underline,
        TextStyle,
        Color,
        FontFamily,
        FontSize,
        LetterSpacing,
        TextAlign.configure({ types: ['heading', 'paragraph'] })
      ],
      content: '<p style="text-align:center;color:#999">点击上方「加载红头模板」开始</p>',
      onSelectionUpdate: ({ editor }) => {
        this.hasSelection = editor.state.selection.from !== editor.state.selection.to
      },
      // onTransaction 每次状态变更都触发(包括 mark 变更),onSelectionUpdate 只在选区范围变化时触发
      onTransaction: ({ editor }) => {
        const { from, to, empty } = editor.state.selection

        // 1. 先尝试从 textStyle mark 读取
        let textStyleMark = null
        if (empty) {
          const $from = editor.state.doc.resolve(from)
          textStyleMark = $from.marks().find(m => m.type.name === 'textStyle')
        } else {
          // 选区可能跨了多个节点,遍历整个选区范围找 mark
          editor.state.doc.nodesBetween(from, to, (node) => {
            if (!textStyleMark && node.marks?.length) {
              textStyleMark = node.marks.find(m => m.type.name === 'textStyle')
            }
          })
        }

        if (textStyleMark) {
          this.currentFontFamily = textStyleMark.attrs.fontFamily || ''
          this.currentFontSize = textStyleMark.attrs.fontSize || ''
          this.currentColor = textStyleMark.attrs.color || '#000000'
        } else {
          // 2. 没有 textStyle mark(样式来自 CSS) — 读取 DOM 计算样式
          this.syncFromComputedStyle(editor, from)
        }
      }
    })
  },

  beforeDestroy () {
    this.editor && this.editor.destroy()
  },

  methods: {
    /* ---------- 模板 ---------- */
    loadTemplate () {
      // 红头公文模板 — 仅红头部分用 inline style 锁死,正文样式交给 .paper 全局 CSS
      const html = `
        <p style="text-align:center" data-govt-title>
          <strong><span style="color:#c00000;font-family:'方正小标宋简体','方正小标宋_GBK',SimSun,serif;font-size:54px;letter-spacing:14px">中华人民共和国应急管理部</span></strong>
        </p>
        <hr>
        <hr>

        <h1>金华市金水区地震应急预案</h1>

        <h2>1 总则</h2>
        <h3>1.1 编制目的</h3>
        <p>贯彻落实习近平总书记关于防范化解重大安全风险和防灾减灾救灾重要论述,依法科学统一、有力有序有效实施地震应急工作,最大程度减少人员伤亡和经济损失,维护社会正常秩序。</p>

        <h3>1.2 编制依据</h3>
        <p>依据《中华人民共和国防震减灾法》《中华人民共和国突发事件应对法》《破坏性地震应急条例》《浙江省防震减灾条例》《浙江省地震应急预案》《金华市地震应急预案》《金华市金东区突发事件总体应急预案》等。</p>

        <h3>1.3 适用范围</h3>
        <p>本预案适用于本区行政区域内发生地震灾害和区外发生对我区造成重大影响的地震灾害的应对工作及其他相关地震事件的应急处置。</p>

        <h3>1.4 工作原则</h3>
        <p>抗震救灾工作坚持党委政府统一领导、部门协同、军地联动、分级负责、属地为主,快速反应、资源共享的工作原则。</p>
      `
      this.editor.commands.setContent(html)
    },

    /* ---------- AI 生成(末尾追加) ---------- */
    async aiGenerate () {
      // 这里换成你的真实 AI 接口
      // const res = await fetch('/api/generate', ...)
      const generated = `
        <h2 style="font-family:'黑体',SimHei,sans-serif;font-size:22px">2 应急组织机构</h2>
        <h3 style="font-family:'楷体',KaiTi,serif;font-size:18px">2.1 指挥部</h3>
        <p style="font-family:'仿宋_GB2312',FangSong,serif;font-size:18px;line-height:1.8;text-indent:2em">
          成立国家地震应急指挥部,由应急管理部部长任总指挥,统一领导、指挥和协调全国地震应急工作。
        </p>
      `
      this.editor.chain().focus('end').insertContent(generated).run()
    },

    /* ---------- 选中文字 AI 润色 ---------- */
    async polishSelection () {
      const { from, to } = this.editor.state.selection
      const selectedText = this.editor.state.doc.textBetween(from, to, ' ')
      if (!selectedText) return

      this.polishing = true
      try {
        // 这里换成你的真实润色接口
        // const res = await fetch('/api/polish', { method:'POST', body: JSON.stringify({ text: selectedText }) })
        // const { polished } = await res.json()
        const polished = await this.mockPolish(selectedText)

        // 把选区替换为润色后的内容,保留原样式
        this.editor.chain().focus().insertContentAt({ from, to }, polished).run()
      } finally {
        this.polishing = false
      }
    },

    mockPolish (text) {
      return new Promise(resolve => {
        setTimeout(() => resolve(`【润色后】${text}(更加正式、规范的表达)`), 600)
      })
    },

    /* ---------- 工具栏样式同步 ---------- */
    syncFromComputedStyle (editor, pos) {
      try {
        const domPos = editor.view.domAtPos(pos)
        // domAtPos 可能返回文本节点或元素节点
        const el = domPos.node.nodeType === 3
          ? domPos.node.parentElement
          : domPos.node instanceof Element ? domPos.node : null
        if (!el) return

        const cs = window.getComputedStyle(el)
        this.currentColor = cs.color || '#000000'
        this.currentFontSize = cs.fontSize || ''

        // fontFamily 特殊处理:getComputedStyle 返回实际渲染字体(如 "SimHei"),
        // 而下拉选项是 CSS font-stack(如 "SimHei, sans-serif"),
        // 遍历下拉选项做模糊匹配
        const raw = cs.fontFamily.replace(/['"]/g, '').toLowerCase()
        const opts = [
          '方正小标宋简体,方正小标宋_GBK,SimSun,serif',
          '仿宋_GB2312, FangSong, serif',
          'SimSun, serif',
          'SimHei, sans-serif',
          'KaiTi, serif'
        ]
        const match = opts.find(o => {
          const first = o.split(',')[0].trim().toLowerCase()
          return raw.includes(first) || first.includes(raw)
        })
        this.currentFontFamily = match || raw
      } catch (_) { /* domAtPos 偶有边界情况 */ }
    },
    setFontFamily (v) {
      v ? this.editor.chain().focus().setFontFamily(v).run()
        : this.editor.chain().focus().unsetFontFamily().run()
    },
    setFontSize (v) {
      v ? this.editor.chain().focus().setFontSize(v).run()
        : this.editor.chain().focus().unsetFontSize().run()
    },
    setColor (v) {
      this.editor.chain().focus().setColor(v).run()
    },

    /* ---------- 导出 ---------- */
    async exportWord () {
      const inner = this.editor.getHTML()
      const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
        body { font-family: '仿宋_GB2312', FangSong, serif; font-size: 18px; line-height: 1.8; }
        h1,h2,h3 { margin: 12px 0; }
        p { margin: 6px 0; }
        hr { border:none; border-top: 2px solid #c00000; }
      </style></head><body>${inner}</body></html>`
      const blob = await htmlDocx.asBlob(html)
      saveAs(blob, 'demo.docx')
    },

    exportMarkdown () {
      // 简易转换:可换成 turndown
      const blob = new Blob([this.editor.getText()], { type: 'text/markdown;charset=utf-8' })
      saveAs(blob, 'demo.md')
    }
  }
}
</script>

<style>
.app { max-width: 1100px; margin: 0 auto; padding: 16px; background: #e8eaed; min-height: 100vh; }

.toolbar-top, .toolbar-format {
  display: flex; flex-wrap: wrap; gap: 6px; align-items: center;
  padding: 8px; background: #f5f5f5; border-radius: 4px; margin-bottom: 8px;
}
.toolbar-format button {
  min-width: 32px; padding: 4px 8px; border: 1px solid #ddd; background: #fff; cursor: pointer;
}
.toolbar-format button.active { background: #1976d2; color: #fff; }
.toolbar-format .divider { width: 1px; height: 20px; background: #ccc; margin: 0 6px; }
.toolbar-format .ai-polish {
  background: linear-gradient(90deg, #7b2ff7, #f107a3); color: #fff; border: none; padding: 6px 12px;
}

/* ------------------------------------------------------------------
 * A4 纸张 + 公文版式 (参照 GB/T 9704-2012)
 * 上 37mm / 下 35mm / 左 28mm / 右 26mm
 * 正文: 仿宋_GB2312 三号 (≈ 21px), 行距 28磅 (≈ 1.6), 首行缩进 2字符
 * ------------------------------------------------------------------ */
.paper {
  width: 210mm;
  min-height: 297mm;
  padding: 37mm 26mm 35mm 28mm;
  background: #fff;
  box-shadow: 0 2px 12px rgba(0,0,0,0.18);
  margin: 16px auto;
  box-sizing: border-box;
}
.paper .ProseMirror {
  outline: none;
  min-height: 220mm;
  font-family: '仿宋_GB2312', '仿宋', FangSong, serif;
  font-size: 21px;
  line-height: 1.6;
  color: #000;
}

/* 正文段落 — 首行缩进 2 字符 */
.paper .ProseMirror p {
  margin: 0;
  text-indent: 2em;
}

/* 公文标题 (二号 方正小标宋) */
.paper .ProseMirror h1 {
  text-align: center;
  font-family: '方正小标宋简体', '方正小标宋_GBK', SimSun, serif;
  font-size: 30px;
  font-weight: bold;
  line-height: 1.5;
  margin: 28px 0 32px;
  text-indent: 0;
}

/* 一级标题 "1 总则" — 黑体 三号 */
.paper .ProseMirror h2 {
  font-family: SimHei, '黑体', sans-serif;
  font-size: 22px;
  font-weight: normal;
  line-height: 1.5;
  margin: 18px 0 4px;
  text-indent: 0;
}

/* 二级标题 "1.1 编制目的" — 楷体 三号 */
.paper .ProseMirror h3 {
  font-family: KaiTi, '楷体_GB2312', '楷体', serif;
  font-size: 21px;
  font-weight: normal;
  line-height: 1.5;
  margin: 10px 0 2px;
  text-indent: 2em;
}

/* 三级标题(如有) — 仿宋加粗 */
.paper .ProseMirror h4 {
  font-family: '仿宋_GB2312', FangSong, serif;
  font-size: 21px;
  font-weight: bold;
  margin: 8px 0 0;
  text-indent: 2em;
}

/* 红色反线 (武文头) — 紧挨的两条 <hr> 自动:细+粗 */
.paper .ProseMirror hr {
  border: none;
  border-top: 1px solid #c00000;
  margin: 6px 0 0;
}
.paper .ProseMirror hr + hr {
  border-top-width: 4px;
  margin: 4px 0 32px;
}
</style>
