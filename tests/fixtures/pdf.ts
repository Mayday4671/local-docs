/** Small, synthetic two-page PDF, including a CJK font that needs bundled CMaps. */
export function pdfFixture(): Buffer {
  const chinese = Buffer.from('离线预览测试', 'utf16le').swap16().toString('hex')
  const streams = [
    'BT /F1 24 Tf 40 240 Td (Offline PDF - page one) Tj ET',
    `BT /F2 24 Tf 40 240 Td <${chinese}> Tj ET`,
  ]
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 500 320] /Resources << /Font << /F1 5 0 R >> >> /Contents 8 0 R >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 500 320] /Resources << /Font << /F2 6 0 R >> >> /Contents 9 0 R >>',
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
    '<< /Type /Font /Subtype /Type0 /BaseFont /STSong-Light /Encoding /UniGB-UCS2-H /DescendantFonts [7 0 R] >>',
    '<< /Type /Font /Subtype /CIDFontType0 /BaseFont /STSong-Light /CIDSystemInfo << /Registry (Adobe) /Ordering (GB1) /Supplement 0 >> /DW 1000 >>',
    ...streams.map(
      (text) => `<< /Length ${Buffer.byteLength(text)} >>\nstream\n${text}\nendstream`,
    ),
  ]
  let content = '%PDF-1.4\n'
  const offsets = [0]
  objects.forEach((object, i) => {
    offsets.push(Buffer.byteLength(content))
    content += `${i + 1} 0 obj\n${object}\nendobj\n`
  })
  const xref = Buffer.byteLength(content)
  content += `xref\n0 ${offsets.length}\n0000000000 65535 f \n`
  content += offsets
    .slice(1)
    .map((offset) => `${String(offset).padStart(10, '0')} 00000 n \n`)
    .join('')
  content += `trailer\n<< /Size ${offsets.length} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF`
  return Buffer.from(content)
}

/** Real uncompressed RGB images on each page exercise large-file range reads. */
export function largePdfFixture(pages = 48): Buffer {
  const objects: (string | Buffer)[] = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    `<< /Type /Pages /Kids [${Array.from({ length: pages }, (_, i) => `${i + 3} 0 R`).join(' ')}] /Count ${pages} >>`,
  ]
  const font = pages + 3
  for (let i = 0; i < pages; i++) {
    const landscape = i % 7 === 6
    objects.push(
      `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${landscape ? '800 600' : '600 800'}] ${i % 11 === 10 ? '/Rotate 90' : ''} /Resources << /Font << /F1 ${font} 0 R >> /XObject << /Im ${font + 2 + i * 2} 0 R >> >> /Contents ${font + 1 + i * 2} 0 R >>`,
    )
  }
  objects.push('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>')
  for (let i = 0; i < pages; i++) {
    const stream = `BT /F1 24 Tf 40 540 Td (Large PDF - page ${i + 1}) Tj ET\nq 400 0 0 400 40 80 cm /Im Do Q`
    objects.push(`<< /Length ${Buffer.byteLength(stream)} >>\nstream\n${stream}\nendstream`)
    const pixels = Buffer.alloc(512 * 512 * 3)
    for (let k = 0; k < pixels.length; k += 3) {
      pixels[k] = 60 + i * 3
      pixels[k + 1] = 130
      pixels[k + 2] = 200
    }
    objects.push(
      Buffer.concat([
        Buffer.from(
          `<< /Type /XObject /Subtype /Image /Width 512 /Height 512 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length ${pixels.length} >>\nstream\n`,
        ),
        pixels,
        Buffer.from('\nendstream'),
      ]),
    )
  }
  const parts = [Buffer.from('%PDF-1.4\n')],
    offsets = [0]
  let length = parts[0].length
  objects.forEach((object, i) => {
    offsets.push(length)
    const bytes = Buffer.concat([
      Buffer.from(`${i + 1} 0 obj\n`),
      Buffer.from(object),
      Buffer.from('\nendobj\n'),
    ])
    parts.push(bytes)
    length += bytes.length
  })
  parts.push(
    Buffer.from(
      `xref\n0 ${offsets.length}\n0000000000 65535 f \n${offsets
        .slice(1)
        .map((offset) => `${String(offset).padStart(10, '0')} 00000 n \n`)
        .join('')}trailer\n<< /Size ${offsets.length} /Root 1 0 R >>\nstartxref\n${length}\n%%EOF`,
    ),
  )
  return Buffer.concat(parts)
}
