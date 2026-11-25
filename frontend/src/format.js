/**
 * 按 Excel「格式要求」列严格格式化数值展示。
 * 格式串示例：'千分位；两位小数' / '百分比；' / '百分比；两位小数'
 * 存储值保持原始数值（小数），仅在展示层格式化，符合"单位转换只在展示层做"。
 */
export function formatValue(value, field) {
  if (value === null || value === undefined || value === '') return null
  const fmt = field?.format || ''
  if (!fmt) return String(value)

  const num = Number(value)
  if (Number.isNaN(num)) return String(value)

  const hasThousands = fmt.includes('千分位')
  const hasPercent = fmt.includes('百分比')
  const hasTwoDecimals = fmt.includes('两位小数')

  if (hasPercent) {
    // 百分比一律强制两位小数：「百分比；」与「百分比；两位小数」显示一致
    return (num * 100).toFixed(2) + '%'
  }

  if (hasThousands) {
    return num.toLocaleString('zh-CN', {
      minimumFractionDigits: hasTwoDecimals ? 2 : 0,
      maximumFractionDigits: hasTwoDecimals ? 2 : 6,
    })
  }

  if (hasTwoDecimals) {
    return num.toFixed(2)
  }

  return String(value)
}

/** 供列标题 Tooltip 显示应用的格式规则 */
export function formatHint(field) {
  return field?.format ? `格式：${field.format}` : null
}
