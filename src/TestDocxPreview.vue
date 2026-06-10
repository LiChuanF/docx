<template>
  <div class="app">
    <div class="toolbar-top">
      <strong>测试库：docx-preview</strong>
      <input type="file" accept=".docx" @change="onFileChange">
      <label>
        <input type="checkbox" v-model="editable" @change="syncEditable"> 让渲染结果可编辑
      </label>
      <span class="hint">{{ status }}</span>
    </div>

    <!-- 原生 selection + execCommand 工具栏 — 仅在可编辑时显示 -->
    <div v-if="editable" class="toolbar-format">
      <select @change="applyFontFamily($event.target.value); $event.target.value = ''">
        <option value="">字体</option>
        <option value="仿宋_GB2312, FangSong, serif">仿宋_GB2312</option>
        <option value="SimSun, serif">宋体</option>
        <option value="SimHei, sans-serif">黑体</option>
        <option value="KaiTi, serif">楷体</option>
        <option value="方正小标宋简体, SimSun, serif">方正小标宋</option>
      </select>

      <select @change="applyFontSize($event.target.value); $event.target.value = ''">
        <option value="">字号</option>
        <option value="42pt">初号</option>
        <option value="36pt">小初</option>
        <option value="26pt">一号</option>
        <option value="24pt">小一</option>
        <option value="22pt">二号</option>
        <option value="18pt">小二</option>
        <option value="16pt">三号</option>
        <option value="15pt">小三</option>
        <option value="14pt">四号</option>
        <option value="12pt">小四</option>
      </select>

      <input type="color" @input="applyColor($event.target.value)">

      <button @mousedown.prevent="exec('bold')"><b>B</b></button>
      <button @mousedown.prevent="exec('italic')"><i>I</i></button>
      <button @mousedown.prevent="exec('underline')"><u>U</u></button>

      <button @mousedown.prevent="applyAlign('left')">左</button>
      <button @mousedown.prevent="applyAlign('center')">中</button>
      <button @mousedown.prevent="applyAlign('right')">右</button>

      <span class="divider" />
      <span class="hint">先选中文字,再点按钮</span>
    </div>

    <div class="paper-wrap">
      <div ref="container" class="docx-container"></div>
    </div>
  </div>
</template>

<script>
import { renderAsync } from 'docx-preview'
import { preprocessShapeLines } from "./utils/preprocessShapeLine";

export default {
  data () {
    return {
      status: '请选择一个 .docx 文件',
      editable: false
    }
  },

  methods: {
    async onFileChange (e) {
      const file = e.target.files[0]
      if (!file) return
      const original = await file.arrayBuffer();
      const buf = await preprocessShapeLines(original);
      this.status = '渲染中…'
      try {
        this.$refs.container.innerHTML = ''
        await renderAsync(buf, this.$refs.container, null, {
          className: 'docx',
          inWrapper: true,
          breakPages: true,
          ignoreLastRenderedPageBreak: true,
          trimXmlDeclaration: true,
          renderHeaders: true,
          renderFooters: true,
          renderFootnotes: true
        })
        this.syncEditable()
        this.status = '渲染成功'
      } catch (err) {
        this.status = '渲染失败：' + err.message
      }
    },

    syncEditable () {
      const wrap = this.$refs.container.querySelector('.docx-wrapper')
      if (wrap) wrap.setAttribute('contenteditable', this.editable ? 'true' : 'false')
    },

    /* ---------- 编辑命令 ----------
     * 思路:
     *   1. B/I/U → document.execCommand 原生支持(已废弃但浏览器仍兼容)
     *   2. 字体/字号/颜色 → 拿当前 Range,把选中范围包一个 <span style="..."> 然后塞回去
     *   3. 对齐 → 找到选区所在的"块级元素"(p/section 等),直接改 style.textAlign
     * 注意: execCommand 标准已废弃,但所有现代浏览器仍然支持 bold/italic/underline,
     *      生产可换成 Range 自己改 DOM(更可控,但代码量大约 3 倍)。
     * ---------------------------- */

    exec (cmd) {
      document.execCommand(cmd, false, null)
    },

    getRange () {
      const sel = window.getSelection()
      if (!sel || sel.rangeCount === 0) return null
      const range = sel.getRangeAt(0)
      // 确认选区在我们的可编辑容器内
      if (!this.$refs.container.contains(range.commonAncestorContainer)) return null
      return range
    },

    wrapSelectionWithStyle (cssText) {
      const range = this.getRange()
      if (!range || range.collapsed) {
        this.status = '请先选中文字'
        return
      }
      const span = document.createElement('span')
      span.setAttribute('style', cssText)
      try {
        // surroundContents 在选区跨越多个节点时会抛错,这里用 extractContents + appendChild 兼容
        const frag = range.extractContents()
        span.appendChild(frag)
        range.insertNode(span)
        // 清掉选区
        window.getSelection().removeAllRanges()
      } catch (err) {
        this.status = '应用样式失败: ' + err.message
      }
    },

    applyFontFamily (font) {
      if (!font) return
      this.wrapSelectionWithStyle(`font-family: ${font}`)
    },

    applyFontSize (size) {
      if (!size) return
      this.wrapSelectionWithStyle(`font-size: ${size}`)
    },

    applyColor (color) {
      this.wrapSelectionWithStyle(`color: ${color}`)
    },

    applyAlign (align) {
      const range = this.getRange()
      if (!range) {
        this.status = '请先把光标放到段落里'
        return
      }
      // 找到光标所在的最近块级元素(p / section / div)
      let node = range.startContainer
      if (node.nodeType === 3) node = node.parentNode
      while (node && node !== this.$refs.container) {
        const display = window.getComputedStyle(node).display
        if (display === 'block' || /^(P|DIV|SECTION|H1|H2|H3|H4|H5|H6|LI)$/.test(node.tagName)) {
          node.style.textAlign = align
          return
        }
        node = node.parentNode
      }
    }
  }
}
</script>

<style>
.app {
  max-width: 1100px;
  margin: 0 auto;
  padding: 16px;
  background: #e8eaed;
  min-height: 100vh;
}

.toolbar-top,
.toolbar-format {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  align-items: center;
  padding: 8px;
  background: #f5f5f5;
  border-radius: 4px;
  margin-bottom: 8px;
}

.toolbar-top .hint,
.toolbar-format .hint {
  color: #666;
  font-size: 13px;
  margin-left: 8px;
}

.toolbar-format button {
  min-width: 32px;
  padding: 4px 8px;
  border: 1px solid #ddd;
  background: #fff;
  cursor: pointer;
}

.toolbar-format .divider {
  width: 1px;
  height: 20px;
  background: #ccc;
  margin: 0 6px;
}

.paper-wrap {
  display: flex;
  justify-content: center;
}

.docx-container :deep(.docx-wrapper) {
  background: #e8eaed;
  padding: 16px 0;
}

.docx-container :deep(.docx-wrapper > section.docx) {
  box-shadow: 0 2px 12px rgba(0, 0, 0, 0.18);
  margin-bottom: 16px;
}

.docx-container :deep([contenteditable="true"]) {
  outline: 1px dashed #1976d2;
}

.docx-container :deep([contenteditable="true"]:focus) {
  outline: 2px solid #1976d2;
}
</style>
