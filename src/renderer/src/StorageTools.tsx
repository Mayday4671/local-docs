import { Archive, Download, FolderOpen, RotateCcw } from 'lucide-react'
import type {
  BackupPreview,
  OperationStatus,
  TransferResult,
  StorageMovePreview,
} from '../../shared/types'
import { fileSize } from './LibraryWorkspace'

export function StorageActions({
  onBackup,
  onRestore,
  onExport,
  result,
}: {
  onBackup: () => void
  onRestore: () => void
  onExport: () => void
  result: TransferResult | null
}) {
  return (
    <>
      <div className="storage-actions">
        <div>
          <Archive size={20} />
          <div>
            <strong>完整备份</strong>
            <p>保留文件、分类、收藏、标记、图片附件、恢复草稿、历史版本与主题。</p>
          </div>
          <button className="secondary" onClick={onBackup}>
            创建备份
          </button>
        </div>
        <div>
          <RotateCcw size={20} />
          <div>
            <strong>从备份恢复</strong>
            <p>先校验、再确认。恢复前自动备份当前文档库。</p>
          </div>
          <button className="secondary" onClick={onRestore}>
            选择备份
          </button>
        </div>
        <div>
          <Download size={20} />
          <div>
            <strong>按分类导出</strong>
            <p>导出当前文件和图片附件；标记、草稿与历史请用完整备份保留。</p>
          </div>
          <button className="secondary" onClick={onExport}>
            导出整个文档库
          </button>
        </div>
      </div>
      {result && (
        <div className="transfer-result" role="status">
          <strong>
            {result.safetyBackupPath
              ? '恢复完成 · 已保留恢复前的备份'
              : `操作完成 · ${result.files} 份文件`}
          </strong>
          <p>{result.safetyBackupPath || result.path}</p>
          <button
            className="secondary"
            onClick={() => void window.localDocs?.revealTransferResult()}
          >
            <FolderOpen size={16} />
            在文件夹中查看
          </button>
        </div>
      )}
    </>
  )
}

export function TransferDialogs({
  active,
  progress,
  preview,
  onCancel,
  onDismissPreview,
  onRestore,
}: {
  active: boolean
  progress: OperationStatus | null
  preview: BackupPreview | null
  onCancel: () => void
  onDismissPreview: () => void
  onRestore: () => void
}) {
  if (active)
    return (
      <div className="modal-backdrop transfer-backdrop">
        <section className="modal" role="dialog" aria-modal="true" aria-labelledby="transfer-title">
          <h2 id="transfer-title">{progress?.phase || '正在准备操作'}</h2>
          <p className="muted">
            {progress?.kind === 'folder-import'
              ? '取消后会保留已导入的文件，再次导入可跳过相同内容。'
              : progress?.kind === 'folder-scan'
                ? '扫描完成后可确认目录和文件，再开始导入。'
                : '请保持应用打开，完成后会显示保存位置。'}
          </p>
          <progress
            aria-label="操作进度"
            max={progress?.total || 1}
            value={progress?.total ? progress.completed : undefined}
          />
          <p aria-live="polite">
            {progress?.total ? `${progress.completed} / ${progress.total}` : '正在处理…'}
          </p>
          <div className="modal-actions">
            <button className="secondary" disabled={!progress?.cancellable} onClick={onCancel}>
              取消操作
            </button>
          </div>
        </section>
      </div>
    )
  if (!preview) return null
  return (
    <div className="modal-backdrop transfer-backdrop">
      <section
        className="modal restore-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="restore-title"
      >
        <h2 id="restore-title">确认恢复文档库</h2>
        <p>已校验这份备份的文件内容。</p>
        <dl className="restore-summary">
          <div>
            <dt>备份时间</dt>
            <dd>{new Date(preview.createdAt).toLocaleString('zh-CN')}</dd>
          </div>
          <div>
            <dt>文件 / 回收站</dt>
            <dd>
              {preview.files} / {preview.trash} 份
            </dd>
          </div>
          <div>
            <dt>分类 / 历史版本</dt>
            <dd>
              {preview.categories} / {preview.versions}
            </dd>
          </div>
          <div>
            <dt>文件与历史内容</dt>
            <dd>{fileSize(preview.bytes)}</dd>
          </div>
        </dl>
        <p className="restore-warning">
          恢复会替换当前文档库，不会与现有内容合并。分类、标签、收藏和主题也会回到备份时的状态。
        </p>
        <p className="muted">
          恢复前会自动保存当前文档库的完整备份。自动备份失败时，不会继续恢复。
        </p>
        <div className="modal-actions">
          <button autoFocus className="secondary" onClick={onDismissPreview}>
            取消
          </button>
          <button className="primary" onClick={onRestore}>
            备份当前资料并恢复
          </button>
        </div>
      </section>
    </div>
  )
}

export function StorageMoveDialog({
  preview,
  onCancel,
  onMove,
}: {
  preview: StorageMovePreview
  onCancel: () => void
  onMove: () => void
}) {
  return (
    <div className="modal-backdrop transfer-backdrop">
      <section
        className="modal storage-move-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="storage-move-title"
      >
        <h2 id="storage-move-title">修改文档库位置</h2>
        <p>把当前文档库迁移到新文件夹，完成后立即生效。</p>
        <dl className="storage-paths">
          <dt>当前位置</dt>
          <dd>{preview.source}</dd>
          <dt>新位置</dt>
          <dd>{preview.target}</dd>
        </dl>
        <p>
          {preview.files} 份文件（含回收站） · {preview.versions} 个历史版本 ·{' '}
          {fileSize(preview.bytes)}
        </p>
        <p className="muted">
          文件、分类、收藏、标记、图片附件、恢复草稿与历史版本会一起复制并校验。旧位置的副本会保留，不再自动同步；取消或失败时继续使用原位置。
        </p>
        <div className="modal-actions">
          <button autoFocus className="secondary" onClick={onCancel}>
            取消
          </button>
          <button className="primary" onClick={onMove}>
            迁移并使用此位置
          </button>
        </div>
      </section>
    </div>
  )
}
