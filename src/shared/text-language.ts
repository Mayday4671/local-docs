export interface TextLanguage {
  language: string | null
  label: string
  source: 'extension' | 'content' | 'plain'
}
export const MAX_SYNTAX_LENGTH = 262_144

// Language detection is separate from byte decoding: Chinese strings stay unchanged.
const languages: Record<string, [string | null, string]> = {
  '.sql': ['sql', 'SQL'],
  '.json': ['json', 'JSON'],
  '.jsonl': ['json', 'JSON Lines'],
  '.xml': ['xml', 'XML'],
  '.html': ['xml', 'HTML'],
  '.htm': ['xml', 'HTML'],
  '.vue': ['xml', 'Vue'],
  '.svelte': ['xml', 'Svelte'],
  '.yaml': ['yaml', 'YAML'],
  '.yml': ['yaml', 'YAML'],
  '.ini': ['ini', 'INI'],
  '.conf': ['ini', '配置文件'],
  '.cfg': ['ini', '配置文件'],
  '.properties': ['properties', 'Properties'],
  '.toml': ['ini', 'TOML'],
  '.env': ['ini', '环境变量'],
  '.css': ['css', 'CSS'],
  '.scss': ['scss', 'SCSS'],
  '.less': ['less', 'Less'],
  '.js': ['javascript', 'JavaScript'],
  '.jsx': ['javascript', 'JSX'],
  '.mjs': ['javascript', 'JavaScript'],
  '.cjs': ['javascript', 'JavaScript'],
  '.ts': ['typescript', 'TypeScript'],
  '.tsx': ['typescript', 'TSX'],
  '.py': ['python', 'Python'],
  '.java': ['java', 'Java'],
  '.c': ['c', 'C'],
  '.h': ['c', 'C'],
  '.cpp': ['cpp', 'C++'],
  '.cc': ['cpp', 'C++'],
  '.hpp': ['cpp', 'C++'],
  '.cs': ['csharp', 'C#'],
  '.go': ['go', 'Go'],
  '.rs': ['rust', 'Rust'],
  '.rb': ['ruby', 'Ruby'],
  '.php': ['php', 'PHP'],
  '.sh': ['bash', 'Shell'],
  '.bat': ['dos', 'Batch'],
  '.cmd': ['dos', 'Batch'],
  '.ps1': ['powershell', 'PowerShell'],
  '.tex': ['latex', 'LaTeX'],
  '.csv': [null, 'CSV'],
  '.tsv': [null, 'TSV'],
  '.log': [null, '日志'],
  '.rst': [null, 'reStructuredText'],
  '.gitignore': [null, 'Git ignore'],
  '.dockerignore': [null, 'Docker ignore'],
}

export function textLanguage(extension: string, text: string): TextLanguage {
  const known = languages[extension.toLowerCase()]
  if (known) return { language: known[0], label: known[1], source: 'extension' }
  // Infer only distinctive structures in generic text, rather than guessing prose as code.
  const sample = text.slice(0, 32_768).trimStart()
  if (/^[{\[]/.test(sample) && text.length <= MAX_SYNTAX_LENGTH) {
    try {
      const value: unknown = JSON.parse(text)
      if (value !== null && typeof value === 'object')
        return { language: 'json', label: 'JSON', source: 'content' }
    } catch {
      // Incomplete JSON is still highlighted when its extension explicitly says JSON.
    }
  }
  const sql = sample.replace(/\/\*[\s\S]*?\*\/|--[^\r\n]*/g, '').trimStart()
  if (
    /^(?:SELECT\s+(?:DISTINCT\s+)?(?:\*|[\w"`\[\].]+(?:\s*,\s*[\w"`\[\].]+)*)\s+FROM\s+[\w"`\[]|INSERT\s+INTO\s+[\w"`\[]|UPDATE\s+[\w"`\[][\s\S]{1,1000}?\bSET\b|DELETE\s+FROM\s+[\w"`\[]|CREATE\s+(?:TABLE|VIEW|INDEX|DATABASE|PROCEDURE)\s+[\w"`\[]|SELECT\s+(?:\d+(?:\.\d+)?|'[^']*')\s*;)/i.test(
      sql,
    )
  )
    return { language: 'sql', label: 'SQL', source: 'content' }
  if (/^<\?xml\s/i.test(sample) || /^<!doctype\s+html\b/i.test(sample))
    return {
      language: 'xml',
      label: /^<!doctype/i.test(sample) ? 'HTML' : 'XML',
      source: 'content',
    }
  const shebang = sample.split(/\r?\n/, 1)[0]
  if (/^#!.*\bpython[\d.]*\b/.test(shebang))
    return { language: 'python', label: 'Python', source: 'content' }
  if (/^#!.*\b(?:ba|z|da)?sh\b/.test(shebang))
    return { language: 'bash', label: 'Shell', source: 'content' }
  if (/^#!.*\bnode\b/.test(shebang))
    return { language: 'javascript', label: 'JavaScript', source: 'content' }
  return { language: null, label: '普通文本', source: 'plain' }
}

export interface SyntaxToken {
  text: string
  classes: string[]
}
