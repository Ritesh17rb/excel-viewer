import { WorkbookViewer } from '@/components/workbook-viewer'
import { getBlobAccess, hasBlobStore } from '@/lib/server/workbook-store'

export default function Page() {
  return (
    <WorkbookViewer
      uploadMode={hasBlobStore() ? 'blob' : 'local'}
      blobAccess={getBlobAccess()}
    />
  )
}

