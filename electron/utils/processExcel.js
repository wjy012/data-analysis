import ExcelJS from 'exceljs'

// Excel序列日期转JS日期
function excelDateToJSDate(serial) {
  return new Date((serial - 25569) * 86400 * 1000)
}

// 统一解析日期
function parseDate(value) {
  if (!value) return null

  if (value instanceof Date) {
    return value
  }

  if (typeof value === 'number') {
    return excelDateToJSDate(value)
  }

  const date = new Date(value)

  return isNaN(date.getTime())
    ? null
    : date
}

// 只按日期计算天数
function calcDays(assignTime, submitTime) {
  const start = parseDate(assignTime)
  const end = parseDate(submitTime)

  if (!start || !end) {
    return ''
  }

  const startDate = new Date(
    start.getFullYear(),
    start.getMonth(),
    start.getDate()
  )

  const endDate = new Date(
    end.getFullYear(),
    end.getMonth(),
    end.getDate()
  )

  return Math.floor(
    (endDate - startDate) / 86400000
  )
}

export async function processSrmData(buffer) {
  const workbook = new ExcelJS.Workbook()

  await workbook.xlsx.load(buffer)

  const sourceSheet = workbook.getWorksheet(1)

  // =========================
  // 第二行中文表头
  // =========================
  const headerRow = sourceSheet.getRow(2)

  const headers = []

  headerRow.eachCell((cell) => {
    headers.push(String(cell.value || '').trim())
  })

  // 找到采购方式位置
  const purchaseMethodIndex =
    headers.indexOf('采购方式')

  if (purchaseMethodIndex === -1) {
    throw new Error('未找到采购方式列')
  }

  // =========================
  // 新表头
  // =========================
  const newHeaders = [...headers]

  // 找到相关列位置
  const materialCodeIndex = newHeaders.indexOf('订单物料代码')

  const materialDescIndex = newHeaders.indexOf('订单物料描述')

  // 删除原位置
  const materialCode = newHeaders.splice(materialCodeIndex, 1)[0]

  const materialDesc = newHeaders.splice(newHeaders.indexOf('订单物料描述'),1)[0]

  // 在采购方式后插入
  const purchaseMethodPos = newHeaders.indexOf('采购方式')

  newHeaders.splice(
    purchaseMethodPos + 1,
    0,
    '订单月份',
    '国内国外',
    '筛选删除',
    '筛选删除（重点服务对象）',
    '采购组织更新',
    materialCode,
    materialDesc
  )

  // 公司名称后新增列
  const companyNameIndex = newHeaders.indexOf('公司名称')

  if (companyNameIndex !== -1) {
    newHeaders.splice(
      companyNameIndex + 1,
      0,
      '修改公司代码',
      '修改公司名称'
    )
  }

  // =========================
  // 转JSON
  // =========================
  const resultData = []

  for (
    let rowIndex = 3;
    rowIndex <= sourceSheet.rowCount;
    rowIndex++
  ) {
    const row = sourceSheet.getRow(rowIndex)

    const values = row.values

    // 空行跳过
    if (!values || values.length <= 1) {
      continue
    }

    const rowData = {}

    headers.forEach((header, index) => {
      rowData[header] = values[index + 1]
    })

    // =====================
    // 国内国外
    // =====================
    rowData['国内国外'] = rowData['是否境外'] === '否'? '国内' : '国外'

    // =====================
    // 订单月份
    // =====================
    let orderMonth = ''

    if (rowData['订单审核时间']) {
      const dateStr = String(
        rowData['订单审核时间']
      ).substring(0, 10)

      const month = Number(dateStr.substring(5, 7))

      const day = Number(dateStr.substring(8, 10))

      orderMonth = day >= 26? (month === 12 ? 1 : month + 1): month
    }

    rowData['订单月份'] = orderMonth

    // =====================
    // 采购组织更新
    // =====================
    const purchaseOrg = Number(rowData['采购组织'])

    const companyName = rowData['买方公司名称']

    let purchaseOrgNew = ''

    if (purchaseOrg === 7129) {
      purchaseOrgNew = companyName === '紫金矿业物流（厦门）有限公司'? 7107: 7129
    } else if (purchaseOrg === 7107) {
      purchaseOrgNew = companyName === '紫金矿业物流有限公司'? 7129: 7107
    }

    rowData['采购组织更新'] = purchaseOrgNew

    // =====================
    // 修改公司代码、修改公司名称
    // =====================
    if (String(rowData['公司代码']) === '7129') {
      rowData['修改公司代码'] = rowData['代理申报公司代码']
      rowData['修改公司名称'] = rowData['代理申报公司']
    } else {
      rowData['修改公司代码'] = rowData['公司代码']
      rowData['修改公司名称'] = rowData['公司名称']
    }

    // =====================
    // 筛选删除
    // =====================
    rowData['筛选删除'] =
      rowData['框架合同类型'] === '年度框架合同' ||
      rowData['框架合同类型'] === '紫金商城' ||
      rowData['订单删除标识'] === '是'
        ? '是' : '否'
    // =====================
    // 筛选删除（重点服务对象）
    // =====================
    rowData['筛选删除（重点服务对象）'] =
      String(rowData['修改公司代码']) === '7129' &&
      rowData['日常/年度'] === '年度'
        ? '是': rowData['筛选删除']

    resultData.push(rowData)
  }

  console.log(`处理完成，共 ${resultData.length} 行`)

  // =========================
  // 创建新Workbook
  // =========================
  const outputWorkbook = new ExcelJS.Workbook()

  const outputSheet = outputWorkbook.addWorksheet('Sheet1')

  // 中文表头
  outputSheet.addRow(newHeaders)

  // 数据
  for (const item of resultData) {
    const row = []

    headers.forEach((header) => {
      // 跳过原位置的物料列
      if (
        header === '订单物料代码' ||
        header === '订单物料描述'
      ) return

      row.push(item[header])

      if (header === '采购方式') {
        row.push(item['订单月份'])
        row.push(item['国内国外'])
        row.push(item['筛选删除'])
        row.push(item['筛选删除（重点服务对象）'])
        row.push(item['采购组织更新'])

        // 插入物料列
        row.push(item['订单物料代码'])
        row.push(item['订单物料描述'])
      }

      if (header === '公司名称') {
        row.push(item['修改公司代码'])
        row.push(item['修改公司名称'])
      }
    })

    outputSheet.addRow(row)
  }

  // 表头样式
  outputSheet.getRow(1).font = {
    bold: true
  }

  return await outputWorkbook.xlsx.writeBuffer()
}

export async function createDurationReport(
  arrayBuffer,
  {
    durationColumn,
    startColumn,
    endColumn,
    deduplicateColumn = '方案单单号'
  }
) {
  try {
    const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(arrayBuffer)

  const sourceSheet = workbook.getWorksheet(1)

  const headerRow = sourceSheet.getRow(2)

  const headers = []

  headerRow.eachCell(cell => {
    headers.push(
      String(cell.text || cell.value || '').trim()
    )
  })

  const headerMap = {}

  headers.forEach((name, index) => {
    headerMap[name] = index + 1
  })

  const requiredColumns = [
    '需求计划池单号',
    '状态',
    '采购模式',
    deduplicateColumn,
    startColumn,
    endColumn
  ]

  const outputWorkbook = new ExcelJS.Workbook()
  const outputSheet = outputWorkbook.addWorksheet('统计结果')

  outputSheet.addRow([
    ...requiredColumns,
    durationColumn
  ])

  const uniqueSet = new Set()

  const outputRows = []

  const uniqueCol = headerMap[deduplicateColumn]

  const startCol = headerMap[startColumn]

  const endCol = headerMap[endColumn]

  for (
    let rowNum = 3;
    rowNum <= sourceSheet.rowCount;
    rowNum++
  ) {
    const row = sourceSheet.getRow(rowNum)

    const uniqueKey = String(
      row.getCell(uniqueCol).text ||
      row.getCell(uniqueCol).value ||
      ''
    ).trim()

    if (!uniqueKey) continue

    if (uniqueSet.has(uniqueKey)) continue

    const startValue = row.getCell(startCol).value

    const endValue = row.getCell(endCol).value

    if (!startValue || !endValue)
      continue

    uniqueSet.add(uniqueKey)

    const resultRow = []

    for (const col of requiredColumns) {
      resultRow.push(
        row.getCell(headerMap[col]).value
      )
    }

    resultRow.push(
      calcDays(startValue, endValue)
    )

    outputRows.push(resultRow)
  }

  outputRows.forEach(row => {
    outputSheet.addRow(row)
  })

  return await outputWorkbook.xlsx.writeBuffer()

  } catch (error) {
    console.error(error)
  }
  
}

export async function processOrderFrameData(arrayBuffer) {
  // ===== 读取源文件 =====
  const sourceWorkbook = new ExcelJS.Workbook()
  await sourceWorkbook.xlsx.load(arrayBuffer)

  const sourceSheet = sourceWorkbook.getWorksheet(1)

  // pandas header=2
  const headerRowNumber = 3

  // ===== 读取表头 =====
  const headerRow = sourceSheet.getRow(headerRowNumber)

  const headers = []

  for (let col = 1; col <= headerRow.cellCount; col++) {
    headers.push(
      headerRow.getCell(col).text?.trim?.() ??
      String(headerRow.getCell(col).value ?? '')
    )
  }

  // ===== 读取数据 =====
  const rows = []

  for (
    let rowNum = headerRowNumber + 1;
    rowNum <= sourceSheet.rowCount;
    rowNum++
  ) {
    const row = sourceSheet.getRow(rowNum)

    const rowData = []

    for (let col = 1; col <= headers.length; col++) {
      rowData.push(row.getCell(col).value)
    }

    rows.push(rowData)
  }

  // ===== 找到列位置 =====
  const purchaseOrgIdx = headers.indexOf('采购组织名称')
  const contractTypeIdx = headers.indexOf('框架合同类型')

  if (purchaseOrgIdx === -1) {
    throw new Error('未找到列：采购组织名称')
  }

  // ===== 插入新列 =====
  const insertPos = purchaseOrgIdx + 1

  headers.splice(insertPos, 0, '筛选删除')

  // ===== 数据处理 =====
  rows.forEach(row => {
    const purchaseOrg = row[purchaseOrgIdx]
    const contractType = row[contractTypeIdx]

    let value

    if (
      purchaseOrg === '紫金矿业物流有限公司采购组织' ||
      purchaseOrg === '紫金矿业物流（厦门）有限公司采购组织'
    ) {
      value = '是'
    } else if (contractType === '合作协议') {
      value = '是'
    } else {
      value = '否'
    }

    row.splice(insertPos, 0, value)
  })

  // ===== 创建新Workbook（模拟df.to_excel）=====
  const outputWorkbook = new ExcelJS.Workbook()
  const outputSheet = outputWorkbook.addWorksheet('Sheet1')

  // 表头
  outputSheet.addRow(headers)

  // 数据
  rows.forEach(row => {
    outputSheet.addRow(row)
  })

  // 返回Excel Buffer
  return await outputWorkbook.xlsx.writeBuffer()
}
