<script setup lang="ts">
import { ref, h, onMounted, onBeforeUnmount } from "vue";
import { DocxEditor } from "@eigenpal/docx-editor-vue";
import type { DocxEditorRef } from "@eigenpal/docx-editor-vue";
import "@eigenpal/docx-editor-vue/styles.css";
import zhCN from "@eigenpal/docx-editor-i18n/zh-CN";

const buf = ref<ArrayBuffer | null>(null);
const editorRef = ref<DocxEditorRef | null>(null);
const polishing = ref(false);

/** 获取当前选中的文字信息 */
function getSelectedText(): {
  paraId: string;
  selectedText: string;
  paragraphText: string;
} | null {
  if (!editorRef.value) return null;
  const sel = editorRef.value.getSelectionInfo();
  if (!sel || !sel.selectedText || !sel.paraId) return null;
  return {
    paraId: sel.paraId,
    selectedText: sel.selectedText,
    paragraphText: sel.paragraphText,
  };
}

/** AI 润色 —— 占位实现：在选中内容后追加 "123"，以修订模式给出建议 */
function polishSelection() {
  const sel = getSelectedText();
  if (!sel) {
    alert("请先选中要润色的文字");
    return;
  }

  polishing.value = true;
  try {
    editorRef.value!.proposeChange({
      paraId: sel.paraId,
      search: sel.selectedText,
      replaceWith: sel.selectedText + "123",
      author: "AI",
    });
  } finally {
    polishing.value = false;
  }
}

/** 渲染自定义工具栏按钮 */
function renderToolbarExtra() {
  return h("div", { class: "toolbar-extra" }, [
    h(
      "button",
      {
        class: "ai-polish-btn",
        disabled: polishing.value,
        onClick: polishSelection,
        title: "选中文字后点击，AI 将以修订模式给出润色建议",
      },
      polishing.value ? "AI 润色中..." : "✨ AI 润色"
    ),
  ]);
}

/** 监听右键菜单触发的 AI 润色事件 */
function onContextMenuPolish() {
  polishSelection();
}
onMounted(() => window.addEventListener("docx-ai-polish", onContextMenuPolish));
onBeforeUnmount(() =>
  window.removeEventListener("docx-ai-polish", onContextMenuPolish)
);
</script>

<template>
  <div class="editor-shell">
    <DocxEditor
      ref="editorRef"
      :document-buffer="buf"
      :i18n="zhCN"
      mode="editing"
      :toolbar-extra="renderToolbarExtra"
    />
  </div>
</template>

<style scoped>
.editor-shell {
  height: 100vh;
  display: flex;
  flex-direction: column;
}

.toolbar-extra {
  display: flex;
  align-items: center;
  margin-left: 8px;
}

.ai-polish-btn {
  padding: 4px 12px;
  border: 1px solid #8b5cf6;
  border-radius: 4px;
  background: linear-gradient(135deg, #8b5cf6, #6d28d9);
  color: #fff;
  font-size: 13px;
  cursor: pointer;
  white-space: nowrap;
  transition: opacity 0.2s;
}

.ai-polish-btn:hover:not(:disabled) {
  opacity: 0.85;
}

.ai-polish-btn:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
