<template>
  <div class="app">
    <div class="toolbar-top">
      <strong>测试库：mammoth.js</strong>
      <input type="file" accept=".docx" @change="onFileChange">
      <span class="hint">{{ status }}</span>
      <button v-if="rawHtml" @click="showRaw = !showRaw">{{ showRaw ? '隐藏' : '查看' }}原始 HTML</button>
      <button v-if="messages.length" @click="showMsg = !showMsg">{{ showMsg ? '隐藏' : '查看' }}解析消息({{ messages.length }})</button>
    </div>

    <div v-if="editor" class="toolbar-format">
      <button :class="{ active: editor.isActive('bold') }" @click="editor.chain().focus().toggleBold().run()"><b>B</b></button>
      <button :class="{ active: editor.isActive('italic') }" @click="editor.chain().focus().toggleItalic().run()"><i>I</i></button>
      <button :class="{ active: editor.isActive('underline') }" @click="editor.chain().focus().toggleUnderline().run()"><u>U</u></button>
      <button :class="{ active: editor.isActive({ textAlign: 'left' }) }" @click="editor.chain().focus().setTextAlign('left').run()">左</button>
      <button :class="{ active: editor.isActive({ textAlign: 'center' }) }" @click="editor.chain().focus().setTextAlign('center').run()">中</button>
      <button :class="{ active: editor.isActive({ textAlign: 'right' }) }" @click="editor.chain().focus().setTextAlign('right').run()">右</button>
    </div>

    <div class="paper">
      <editor-content :editor="editor" />
    </div>

    <pre v-if="showRaw" class="raw">{{ rawHtml }}</pre>
    <pre v-if="showMsg" class="raw">{{ messages.map(m => `[${m.type}] ${m.message}`).join('\n') }}</pre>
  </div>
</template>

<script>
import { Editor, EditorContent } from '@tiptap/vue-2'
import { Extension } from '@tiptap/core'
import StarterKit from '@tiptap/starter-kit'
import TextAlign from '@tiptap/extension-text-align'
import { TextStyle, Color, FontFamily, FontSize } from '@tiptap/extension-text-style'
import { Underline } from '@tiptap/extension-underline'
import mammoth from 'mammoth'

// tiptap 默认 schema 不允许 paragraph/horizontalRule 节点带任意 class —
// 这里手动给它们加上 class 属性的 parse/render,保留 mammoth styleMap 产出的类名
const PreserveClass = Extension.create({
  name: 'preserveClass',
  addGlobalAttributes () {
    return [{
      types: ['paragraph', 'horizontalRule', 'heading'],
      attributes: {
        class: {
          default: null,
          parseHTML: el => el.getAttribute('class') || null,
          renderHTML: attrs => attrs.class ? { class: attrs.class } : {}
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
      rawHtml: '',
      showRaw: false,
      messages: [],
      showMsg: false,
      status: '请选择一个 .docx 文件'
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
        PreserveClass,
        TextAlign.configure({ types: ['heading', 'paragraph'] })
      ],
      content: '<p style="text-align:center;color:#999">选择 docx 文件后,mammoth 解析的 HTML 会灌进这里</p>'
    })
  },

  beforeDestroy () {
    this.editor && this.editor.destroy()
  },

  methods: {
    async onFileChange (e) {
      const file = e.target.files[0]
      if (!file) return
      this.status = '解析中…'
      try {
        const arrayBuffer = await file.arrayBuffer()
        // convertToHtml 输出语义化标签 (h1/h2/p/strong/em),会丢失字号/字体/颜色等内联样式
        // styleMap 可把 docx 样式名映射到自定义标签,这里先用默认看效果
        const result = await mammoth.convertToHtml(
          { arrayBuffer },
          {
            // styleMap 把 docx 中的命名样式映射成带 class 的标签
            // mammoth 不会输出 docx 段落边框,所以红线那两个段落会变成空 <p>,后处理转成 <hr>
            styleMap: [
              "p[style-name='Govt Red Header'] => p.govt-red-header:fresh",
              "p[style-name='Govt Closing Rule'] => p.govt-closing-rule:fresh",
              "p[style-name='Heading 1'] => h1:fresh",
              "p[style-name='Heading 2'] => h2:fresh",
              "p[style-name='Heading 3'] => h3:fresh"
            ]
          }
        )
        // 后处理:红线是公文"格式",mammoth 会丢段落边框,所以在前端补
        //   - 红头后:粗红线 + 细红线(GB/T 9704 武文头)
        //   - 结尾空段落(styleMap 已转 .govt-closing-rule):转成结尾粗红线
        let html = result.value
        // [^<]* 改成 [\s\S]*? — 防止段落里有 <br>/<strong> 等子标签
        html = html.replace(
          /(<p class="govt-red-header">[\s\S]*?<\/p>)/,
          '$1<hr class="govt-rule-thick" /><hr class="govt-rule-thin" />'
        )
        html = html.replace(
          /<p class="govt-closing-rule">[\s\S]*?<\/p>/g,
          '<hr class="govt-rule-thick" />'
        )
        this.rawHtml = html
        this.messages = result.messages
        const isEmpty = !result.value || result.value.trim() === ''
        if (isEmpty) {
          this.editor.commands.setContent(
            '<p style="text-align:center;color:#c00000">⚠️ mammoth 返回空 HTML — 这个 docx 可能是 html-docx-js 生成的伪 docx(用 altChunk 包了 MHT),不是标准 docx。请用真正的 Word/WPS 另存为或后端用 docxtemplater/officegen 生成。</p>'
          )
        } else {
          this.editor.commands.setContent(html)
        }
        const warnings = result.messages.filter(m => m.type === 'warning').length
        this.status = `解析${isEmpty ? '为空' : '成功'}(${result.messages.length} 条消息,${warnings} 条警告)`
      } catch (err) {
        this.status = '解析失败：' + err.message
      }
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
.toolbar-top .hint { color: #666; font-size: 13px; margin-left: 8px; }
.toolbar-format button {
  min-width: 32px; padding: 4px 8px; border: 1px solid #ddd; background: #fff; cursor: pointer;
}
.toolbar-format button.active { background: #1976d2; color: #fff; }

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
  /* 让 ProseMirror 撑满 paper 内可用高度,再用 flex 把最后一根红线推到底部 */
  min-height: calc(297mm - 37mm - 35mm);
  display: flex;
  flex-direction: column;
  font-family: '仿宋_GB2312', '仿宋', FangSong, serif;
  font-size: 21px;
  line-height: 1.6;
  color: #000;
}
/* 结尾红线之前塞一个 auto margin,把红线挤到 paper 底部 */
.paper .ProseMirror hr.govt-rule-thick:last-child {
  margin-top: auto;
}
.paper .ProseMirror p { margin: 0; text-indent: 2em; }
.paper .ProseMirror h1 {
  text-align: center;
  font-family: '方正小标宋简体', '方正小标宋_GBK', SimSun, serif;
  font-size: 30px; font-weight: bold; line-height: 1.5;
  margin: 28px 0 32px; text-indent: 0;
}
.paper .ProseMirror h2 {
  font-family: SimHei, '黑体', sans-serif;
  font-size: 22px; font-weight: normal; line-height: 1.5;
  margin: 18px 0 4px; text-indent: 0;
}
.paper .ProseMirror h3 {
  font-family: KaiTi, '楷体_GB2312', '楷体', serif;
  font-size: 21px; font-weight: normal; line-height: 1.5;
  margin: 10px 0 2px; text-indent: 2em;
}

/* 红头大字 — 由 styleMap 把 docx 命名样式 "Govt Red Header" 转出来 */
.paper .ProseMirror p.govt-red-header {
  text-align: center;
  color: #c00000;
  font-family: '方正小标宋简体', '方正小标宋_GBK', SimSun, serif;
  font-size: 44px;
  font-weight: bold;
  letter-spacing: 4px;
  text-indent: 0;
  margin: 12px 0 6px;
}

/* tiptap 的 horizontalRule 默认 hr 没边框只有上 margin,这里全部覆盖 */
.paper .ProseMirror hr {
  border: none;
  background: none;
  height: auto;
}
/* 红线 — 武文头:粗(上) + 细(下,紧贴) */
.paper .ProseMirror hr.govt-rule-thick {
  border-top: 4px solid #c00000;
  margin: 4px 0 2px;
}
.paper .ProseMirror hr.govt-rule-thin {
  border-top: 1px solid #c00000;
  margin: 0 0 24px;
}

.raw {
  max-width: 1100px; margin: 16px auto; padding: 12px;
  background: #1e1e1e; color: #d4d4d4; border-radius: 4px;
  font-size: 12px; white-space: pre-wrap; word-break: break-all;
}
</style>
