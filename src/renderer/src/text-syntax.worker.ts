import { syntaxTokens } from './text-syntax'

self.onmessage = (event: MessageEvent<{ language: string; text: string }>) => {
  const { language, text } = event.data
  self.postMessage(syntaxTokens(language, text))
}
