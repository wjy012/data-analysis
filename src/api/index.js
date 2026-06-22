const BASE_URL = 'http://172.22.35.101:8001'

export function exportReport(token, data) {
  console.log(data)
  return window.electronAPI.request({
    url: `${BASE_URL}/zjark-cloud-srmreport/purchaseTrackingReportYesterdayResult/cloudExportExcel`,
    method: 'POST',
    token,
    data
  })
}

export function queryExportTask(token) {
  return window.electronAPI.request({
    url: `${BASE_URL}/zjark-cloud-admin/cloudExport/management/query`,
    method: 'POST',
    token,
    data: {
      blocks: {
        paramBlock: {
          limit: 10,
          offset: 1,
          ascDesc: 'desc',
          orderBy: 'createTime',
          data: {},
          blockId: 'paramBlock'
        }
      }
    }
  })
}

export function downloadFile(url, fileName, taskType) {
  return window.electronAPI.downloadFile({
    url,
    fileName,
    taskType
  })
}