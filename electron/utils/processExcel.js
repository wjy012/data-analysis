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

function calcHours(assignTime, submitTime) {
  const start = parseDate(assignTime)
  const end = parseDate(submitTime)

  if (!start || !end) {
    return ''
  }

  return (end - start) / (1000 * 60 * 60)
}

export async function processSrmData(buffer) {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(buffer)

  const sourceSheet = workbook.getWorksheet(1)

  if (!sourceSheet) {
    throw new Error('未找到第一个工作表')
  }

  // =========================
  // 第二行中文表头
  // =========================
  const headerRow = sourceSheet.getRow(2)
  const headers = []

  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    headers[colNumber - 1] = String(cell.value || '').trim()
  })

  // =========================
  // 必要表头校验
  // =========================
  const requiredHeaders = [
    '采购方式',
    '订单物料代码',
    '订单物料描述',
    '是否境外',
    '代理申报公司代码',
    '采购组织',
    '采购组织名称'
  ]

  const missingHeaders = requiredHeaders.filter(header => !headers.includes(header))

  if (missingHeaders.length > 0) {
    throw new Error(`缺少必要列：${missingHeaders.join('、')}`)
  }

  // =========================
  // 新表头
  // =========================
  const newHeaders = [...headers]

  // 删除订单物料代码原位置
  const materialCodeIndex = newHeaders.indexOf('订单物料代码')
  const materialCode = newHeaders.splice(materialCodeIndex, 1)[0]

  // 删除订单物料描述原位置
  const materialDescIndex = newHeaders.indexOf('订单物料描述')
  const materialDesc = newHeaders.splice(materialDescIndex, 1)[0]

  // 在采购方式后插入新增列和物料列
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
  // 单元格值转文本
  // =========================
  const getCellText = value => {
    if (value === null || value === undefined) {
      return ''
    }

    if (typeof value === 'object') {
      // 公式单元格
      if ('result' in value) {
        return String(value.result ?? '').trim()
      }

      // 富文本单元格
      if (Array.isArray(value.richText)) {
        return value.richText.map(item => item.text || '').join('').trim()
      }

      // 超链接等文本单元格
      if ('text' in value) {
        return String(value.text ?? '').trim()
      }
    }

    return String(value).trim()
  }

  // =========================
  // 日期转换
  // =========================
  const getDateParts = value => {
    if (!value) {
      return null
    }

    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return {
        month: value.getMonth() + 1,
        day: value.getDate()
      }
    }

    const text = getCellText(value)
    const match = text.match(/(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})/)

    if (!match) {
      return null
    }

    return {
      month: Number(match[2]),
      day: Number(match[3])
    }
  }

  // =========================
  // 转JSON并执行数据校验
  // =========================
  const resultData = []

  for (let rowIndex = 3; rowIndex <= sourceSheet.rowCount; rowIndex++) {
    const row = sourceSheet.getRow(rowIndex)
    const values = row.values

    // 空行跳过
    if (!values || values.length <= 1) {
      continue
    }

    const rowData = {}

    headers.forEach((header, index) => {
      if (!header) {
        return
      }

      rowData[header] = values[index + 1]
    })

    // 整行没有有效数据时跳过
    const hasValue = headers.some(header => getCellText(rowData[header]) !== '')

    if (!hasValue) {
      continue
    }

    const overseasFlag = getCellText(rowData['是否境外'])
    const agentCompanyCode = getCellText(rowData['代理申报公司代码'])
    const purchaseOrg = getCellText(rowData['采购组织'])
    const purchaseOrgName = getCellText(rowData['采购组织名称'])

    // =====================
    // 数据校验1
    // 是否境外为“否”且代理申报公司代码有值时，
    // 检查是否包含英文字母
    // =====================
    const invalidOverseas = overseasFlag === '否' &&
      agentCompanyCode !== '' &&
      /[a-zA-Z]/.test(agentCompanyCode)

    // =====================
    // 数据校验2
    // 采购组织与采购组织名称对应关系
    // =====================
    let invalidPurchaseOrgName = false

    if (purchaseOrg === '7129') {
      invalidPurchaseOrgName = purchaseOrgName !== '紫金矿业物流有限公司采购组织'
    } else if (purchaseOrg === '7107') {
      invalidPurchaseOrgName = purchaseOrgName !== '紫金矿业物流（厦门）有限公司采购组织'
    }

    rowData.__validation = {
      invalidOverseas,
      invalidPurchaseOrgName
    }

    // =====================
    // 国内国外
    // =====================
    rowData['国内国外'] = overseasFlag === '否' ? '国内' : '国外'

    // =====================
    // 订单月份
    // =====================
    let orderMonth = ''
    const dateParts = getDateParts(rowData['订单审核时间'])

    if (dateParts) {
      const { month, day } = dateParts
      orderMonth = day >= 26 ? (month === 12 ? 1 : month + 1) : month
    }

    rowData['订单月份'] = orderMonth

    // =====================
    // 采购组织更新
    // =====================
    const purchaseOrgNumber = Number(purchaseOrg)
    const buyerCompanyName = getCellText(rowData['买方公司名称'])
    let purchaseOrgNew = ''

    if (purchaseOrgNumber === 7129) {
      purchaseOrgNew = buyerCompanyName === '紫金矿业物流（厦门）有限公司' ? 7107 : 7129
    } else if (purchaseOrgNumber === 7107) {
      purchaseOrgNew = buyerCompanyName === '紫金矿业物流有限公司' ? 7129 : 7107
    }

    rowData['采购组织更新'] = purchaseOrgNew

    // =====================
    // 修改公司代码、修改公司名称
    // =====================
    if (getCellText(rowData['公司代码']) === '7129' || '7680') {
      rowData['修改公司代码'] = rowData['代理申报公司代码']? rowData['代理申报公司代码']: rowData['公司代码']
      rowData['修改公司名称'] = rowData['代理申报公司']? rowData['代理申报公司']: rowData['公司名称']
    } else {
      rowData['修改公司代码'] = rowData['公司代码']
      rowData['修改公司名称'] = rowData['公司名称']
    }

    // =====================
    // 筛选删除
    // =====================
    const frameworkContractType = getCellText(rowData['框架合同类型'])
    const orderDeleteFlag = getCellText(rowData['订单删除标识'])

    rowData['筛选删除'] =
      frameworkContractType === '年度框架合同' ||
      frameworkContractType === '紫金商城' ||
      orderDeleteFlag === '是'
        ? '是'
        : '否'

    // =====================
    // 筛选删除（重点服务对象）
    // =====================
    rowData['筛选删除（重点服务对象）'] =
      getCellText(rowData['修改公司代码']) === '7129' &&
      getCellText(rowData['日常/年度']) === '年度'
        ? '是'
        : rowData['筛选删除']

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

  // 获取需要高亮的列号
  const overseasColumnNumber = newHeaders.indexOf('是否境外') + 1
  const purchaseOrgNameColumnNumber = newHeaders.indexOf('采购组织名称') + 1

  // 高亮样式：浅红色背景、深红色字体
  const errorFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: {
      argb: 'FFFFC7CE'
    }
  }

  // =========================
  // 写入数据并设置高亮
  // =========================
  for (const item of resultData) {
    const outputRowData = []

    headers.forEach(header => {
      // 跳过物料列原位置
      if (header === '订单物料代码' || header === '订单物料描述') {
        return
      }

      outputRowData.push(item[header])

      if (header === '采购方式') {
        outputRowData.push(item['订单月份'])
        outputRowData.push(item['国内国外'])
        outputRowData.push(item['筛选删除'])
        outputRowData.push(item['筛选删除（重点服务对象）'])
        outputRowData.push(item['采购组织更新'])
        outputRowData.push(item['订单物料代码'])
        outputRowData.push(item['订单物料描述'])
      }

      if (header === '公司名称') {
        outputRowData.push(item['修改公司代码'])
        outputRowData.push(item['修改公司名称'])
      }
    })

    const outputRow = outputSheet.addRow(outputRowData)

    // 校验1不通过时高亮“是否境外”
    if (item.__validation.invalidOverseas && overseasColumnNumber > 0) {
      const cell = outputRow.getCell(overseasColumnNumber)
      cell.fill = errorFill
      cell.font = {
        ...cell.font,
      }
    }

    // 校验2不通过时高亮“采购组织名称”
    if (item.__validation.invalidPurchaseOrgName && purchaseOrgNameColumnNumber > 0) {
      const cell = outputRow.getCell(purchaseOrgNameColumnNumber)
      cell.fill = errorFill
      cell.font = {
        ...cell.font,
      }
    }
  }

  // =========================
  // 表头样式
  // =========================
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
      calcHours(startValue, endValue)
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
