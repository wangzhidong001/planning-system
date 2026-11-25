import React from 'react'
import { Form, Input, InputNumber, Select, DatePicker, Tag, Tooltip } from 'antd'
import dayjs from 'dayjs'
import { formatHint } from './format'

// 只读字段：公式字段（ƒ）与自动生成字段（系统发号）
export function isReadOnly(f) {
  return f.type === 'formula' || f.type === 'auto_id' || (f.auto && f.type !== 'enum')
}

// 按字段元数据生成表单项。field.key 作为表单名（稳定），显示中文 label。
export default function DynamicForm({ fields, initialValues, form, disabledAuto = true }) {
  const normInitial = React.useMemo(() => {
    const v = {}
    fields.forEach(f => {
      let val = initialValues?.[f.key]
      if (f.type === 'date' && val && /^\d{4}-\d{2}-\d{2}$/.test(val)) {
        val = dayjs(val)
      }
      v[f.key] = val
    })
    return v
  }, [initialValues, fields])

  return (
    <Form form={form} layout="vertical" initialValues={normInitial} preserve={false}
      style={{ maxHeight: '60vh', overflowY: 'auto', paddingRight: 8 }}>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0 16px' }}>
        {fields.map(f => {
          const readOnly = disabledAuto && isReadOnly(f)
          const isFormula = f.type === 'formula'
          const isAutoId = f.type === 'auto_id'
          let control
          if (isFormula) {
            control = <Input disabled placeholder="公式字段：由系统按口径自动计算（只读）" />
          } else if (isAutoId) {
            control = <Input disabled placeholder="系统自动发号（生成后不变）" />
          } else if (f.auto && disabledAuto) {
            control = <Input disabled placeholder="系统自动取值（跨表带出）" />
          } else if (f.type === 'enum' && f.enum?.length) {
            control = (
              <Select allowClear placeholder={`请选择${f.name}`}
                options={f.enum.map(e => ({ label: e, value: e }))} />
            )
          } else if (f.type === 'enum' && f.enum_ref) {
            control = <Input placeholder={f.enum_ref.replace(/\n/g, ' ')} />
          } else if (f.type === 'integer' || f.type === 'number') {
            control = <InputNumber style={{ width: '100%' }}
              placeholder={f.format ? `按格式录入：${f.format}` : `请输入${f.name}`}
              thousandSeparator={f.format?.includes('千分位') ? ',' : undefined}
              precision={f.type === 'integer' ? 0 : (f.format?.includes('两位小数') ? 2 : undefined)} />
          } else if (f.type === 'date') {
            control = <DatePicker style={{ width: '100%' }} placeholder={`请选择${f.name}`} />
          } else if (f.type === 'time') {
            control = <Input placeholder="如 2026-06-01 或 2026-06" />
          } else if (f.type === 'url') {
            control = <Input placeholder="请输入链接地址（https://...）" />
          } else {
            control = <Input placeholder={`请输入${f.name}`} />
          }
          return (
            <Form.Item key={f.key} name={f.key} label={
              <Tooltip title={
                <div>
                  {f.desc && <div style={{ marginBottom: 4 }}>{f.desc}</div>}
                  {isFormula && <div>公式：{(f.formula || '').replace(/\n/g, ' ').slice(0, 120)}</div>}
                  {f.format && <div>{formatHint(f)}；百分比字段请录入小数（如 0.15 表示 15%）</div>}
                </div>
              }>
                <span style={{ cursor: 'help' }}>{f.name} {f.required && <span style={{ color: 'red' }}>*</span>}
                  {isFormula && <Tag color="purple" style={{ marginLeft: 6 }}>ƒ公式</Tag>}
                  {isAutoId && <Tag color="blue" style={{ marginLeft: 6 }}>自动发号</Tag>}
                  {f.auto && !isAutoId && !isFormula && <Tag color="blue" style={{ marginLeft: 6 }}>自动</Tag>}
                  {f.format && <Tag color="gold" style={{ marginLeft: 6, fontSize: 10 }}>{f.format}</Tag>}
                </span>
              </Tooltip>
            } rules={[
              ...(f.required && !readOnly ? [{ required: true, message: `请输入${f.name}` }] : []),
              ...(f.validate === '枚举值校验' && f.enum?.length && !readOnly
                ? [{ validator: (_, v) => (!v || f.enum.includes(v)) ? Promise.resolve() : Promise.reject(new Error('枚举值非法')) }] : []),
            ]}>
              {control}
            </Form.Item>
          )
        })}
      </div>
    </Form>
  )
}

// 提交前把 dayjs 转字符串；只读字段不提交
export function serializeValues(fields, values) {
  const out = {}
  fields.forEach(f => {
    if (isReadOnly(f)) return  // 只读字段不提交，由系统维护
    let v = values[f.key]
    if (f.type === 'date' && v && dayjs.isDayjs(v)) v = v.format('YYYY-MM-DD')
    out[f.key] = v ?? null
  })
  return out
}
