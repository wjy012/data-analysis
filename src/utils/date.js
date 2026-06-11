export function getDateRange(type) {
  const now = new Date()

  const year = now.getFullYear()
  const month = now.getMonth()

  if (type === 'year') {
    return [
      `${year - 1}-12-26 00:00:00`,
      `${year}-${String(month + 1).padStart(2, '0')}-25 23:59:59`
    ]
  }

  const prevMonth = month === 0 ? 12 : month
  const prevYear = month === 0 ? year - 1 : year

  return [
    `${prevYear}-${String(prevMonth).padStart(2, '0')}-26 00:00:00`,
    `${year}-${String(month + 1).padStart(2, '0')}-25 23:59:59`
  ]
}