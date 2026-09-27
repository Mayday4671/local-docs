import type { ReactNode } from 'react'
import { ResizableSplit } from './ResizableSplit'

export function EditorSplit({ children }: { children: [ReactNode, ReactNode] }) {
  return (
    <ResizableSplit
      className="editor-split"
      preferenceKey="local-docs:markdown-editor-ratio"
      leftLabel="编辑区"
      rightLabel="预览区"
      resetLabel="均分"
    >
      {children}
    </ResizableSplit>
  )
}
