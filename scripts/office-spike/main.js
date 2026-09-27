import { SuperDoc } from 'superdoc'
import 'superdoc/style.css'
document.body.style.margin = '0'
document.body.style.background = '#edf1f8'
window.spikeErrors = []
window.loadDocument = async (base64) => {
  window.spikeEditor?.destroy()
  document.querySelector('#editor').innerHTML = ''
  window.ready = false
  const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))
  await new Promise((resolve, reject) => {
    window.spikeEditor = new SuperDoc({
      selector: '#editor',
      document: new File([bytes], 'fixture.docx', {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }),
      ui: { toolbar: { container: '#toolbar' } },
      telemetry: { enabled: false },
      onReady: () => {
        window.ready = true
        resolve()
      },
      onException: ({ error }) => {
        window.spikeErrors.push(String(error))
        reject(error)
      },
      onContentError: ({ error }) => {
        window.spikeErrors.push(String(error))
        reject(error)
      },
    })
  })
}
window.exportDocument = async () => {
  const blob = await window.spikeEditor.export({ triggerDownload: false })
  const bytes = new Uint8Array(await blob.arrayBuffer())
  let data = ''
  for (const byte of bytes) data += String.fromCharCode(byte)
  return btoa(data)
}
