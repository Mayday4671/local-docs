import { common, createLowlight } from 'lowlight'
import dos from 'highlight.js/lib/languages/dos'
import latex from 'highlight.js/lib/languages/latex'
import powershell from 'highlight.js/lib/languages/powershell'
import properties from 'highlight.js/lib/languages/properties'
import { MAX_SYNTAX_LENGTH, type SyntaxToken } from '../../shared/text-language'

const syntax = createLowlight({ ...common, dos, latex, powershell, properties })
const MAX_TOKENS = 20_000

/** Return only text and style classes, never HTML, attributes, URLs or executable code. */
export function syntaxTokens(language: string, text: string): SyntaxToken[] | null {
  if (text.length > MAX_SYNTAX_LENGTH || !syntax.registered(language)) return null
  try {
    const tree = syntax.highlight(language, text)
    const tokens: SyntaxToken[] = []
    type Node = (typeof tree.children)[number]
    const visit = (node: Node, classes: string[], depth: number) => {
      if (tokens.length >= MAX_TOKENS || depth > 64) throw new Error('Syntax tree too large')
      if (node.type === 'text') tokens.push({ text: node.value, classes })
      else if (node.type === 'element') {
        const own = Array.isArray(node.properties.className)
          ? node.properties.className.filter((v): v is string => typeof v === 'string')
          : []
        const next = [...classes, ...own.filter((v) => /^[\w-]+$/.test(v))]
        for (const child of node.children) visit(child, next, depth + 1)
      }
    }
    for (const node of tree.children) visit(node, [], 0)
    return tokens.map((token) => token.text).join('') === text ? tokens : null
  } catch {
    return null
  }
}
