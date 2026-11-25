import React, { useEffect, useState } from 'react'
import {
  Table, Button, Input, Space, Modal, Form, message, Tag, Tooltip,
  Upload, Progress, Alert, Popconfirm, Typography
} from 'antd'
import {
  PlusOutlined, EditOutlined, DeleteOutlined, SearchOutlined,
  UploadOutlined, DownloadOutlined, FileExcelOutlined, ReloadOutlined
} from '@ant-design/icons'
import { api } from './api'
import DynamicForm, { serializeValues } from './DynamicForm'
import { formatValue, formatHint } from './format'

const { Text } = Typography

function saveBlob(blob, filename) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a'); a.href = url; a.download = filename; a.click()
  URL.revokeObjectURL(url)
}

export default function DataPage({ moduleCode, moduleName, user }) {
  const [fields, setFields] = useState([])
  const [rows, setRows] = useState([])
  const [total, setTotal] = useState(0)
  const [page, setPage] = useState(1)
  const [pageSize, setPageSize] = useState(20)
  const [fieldQueries, setFieldQueries] = useState({})
  const [loading, setLoading] = useState(false)
  const [editOpen, setEditOpen] = useState(false)
  const [editing, setEditing] = useState(null)
  const [form] = Form.useForm()
  const [importOpen, setImportOpen] = useState(false)
  const [preview, setPreview] = useState(null)
  const [importing, setImporting] = useState(false)

  const can = (b) => user.buttons.includes(b)

  const loadMeta = async () => {
    const m = await api.moduleMeta(moduleCode)
    setFields(m.fields)
  }
  const loadData = async () => {
    setLoading(true)
    try {
      const d = await api.list(moduleCode, '', page, pageSize, fieldQueries)
      setRows(d.rows); setTotal(d.total)
    } catch (e) { message.error(e.message) }
    finally { setLoading(false) }
  }
  useEffect(() => { loadMeta(); setFieldQueries({}); setPage(1) }, [moduleCode])
  useEffect(() => { if (fields.length) loadData() }, [moduleCode, fieldQueries, page, pageSize, fields])

  const openCreate = () => {
    setEditing(null); form.resetFields(); setEditOpen(true)
  }
  const openEdit = (row) => {
    setEditing(row); form.resetFields(); setEditOpen(true)
  }
  const submit = async () => {
    const vals = await form.validateFields()
    const payload = serializeValues(fields, vals)
    try {
      const res = editing ? await api.update(moduleCode, editing._id, payload)
                          : await api.create(moduleCode, payload)
      if (!res.ok) { message.error((res.errors || []).join('；')); return }
      message.success(editing ? '更新成功' : '新增成功')
      setEditOpen(false); loadData()
    } catch (e) { message.error(e.message || '保存失败') }
  }
  const remove = async (ids) => {
    const res = await api.remove(moduleCode, ids)
    if (res.ok) { message.success(`已删除 ${res.deleted} 条`); loadData() }
  }

  // 导入
  const beforeUpload = async (file) => {
    setImporting(true)
    try {
      const p = await api.importPreview(moduleCode, file)
      setPreview(p)
    } catch (e) { message.error(e.message) }
    finally { setImporting(false) }
    return false
  }
  const commitImport = async () => {
    const validRows = preview.rows.filter(r => r.valid).map(r => r.values)
    if (!validRows.length) { message.warning('没有可导入的有效数据'); return }
    const res = await api.importCommit(moduleCode, validRows)
    message.success(`成功导入 ${res.imported} 条`)
    setImportOpen(false); setPreview(null); loadData()
  }

  const queryFields = fields.filter(f => f.query)
  const columns = [
    { title: '序号', dataIndex: '_id', width: 70, fixed: 'left', render: (_, __, i) => (page - 1) * pageSize + i + 1 },
    ...fields.map(f => ({
      title: <Tooltip title={
        <div>
          {f.desc && <div style={{ marginBottom: 4 }}>{f.desc}</div>}
          {f.type === 'formula' && <div>ƒ 公式字段（只读，系统按口径计算）</div>}
          {formatHint(f) && <div>{formatHint(f)}</div>}
        </div>
      }>{f.name + (f.query ? ' 🔍' : '') + (f.type === 'formula' ? ' ƒ' : '')}</Tooltip>,
      dataIndex: f.key,
      ellipsis: true,
      width: f.type === 'date' ? 130 : 160,
      align: (f.type === 'number' || f.type === 'integer') ? 'right' : 'left',
      render: (v, row) => {
        const status = row?.['_fs_' + f.key]
        if (f.type === 'formula' && status && status !== 'OK') {
          return <Tooltip title={`公式状态：${status}`}><Tag color="orange">不适用</Tag></Tooltip>
        }
        if (v === null || v === undefined || v === '') return <Text type="secondary">—</Text>
        const shown = formatValue(v, f)
        if (f.type === 'formula') return <span style={{ color: '#722ed1' }}>{shown}</span>
        if (f.type === 'auto_id' || f.auto) return <Tag color="blue">{shown}</Tag>
        if (f.type === 'url') return <a href={shown} target="_blank" rel="noreferrer">{shown}</a>
        return shown
      }
    })),
    { title: '状态', dataIndex: '_status', width: 90,
      render: v => <Tag color={v === '草稿' ? 'default' : 'green'}>{v}</Tag> },
    ...((can('edit') || can('delete')) ? [{
      title: '操作', key: 'op', width: 130, fixed: 'right',
      render: (_, row) => (
        <Space size={0}>
          {can('edit') && <Button type="link" size="small" icon={<EditOutlined />} onClick={() => openEdit(row)}>编辑</Button>}
          {can('delete') && <Popconfirm title="确认删除？" onConfirm={() => remove([row._id])}>
            <Button type="link" size="small" danger icon={<DeleteOutlined />}>删除</Button>
          </Popconfirm>}
        </Space>
      )
    }] : [])
  ]

  return (
    <div>
      {/* 查询区：每个「是否查询字段」一个独立输入框 */}
      {queryFields.length > 0 && (
        <div style={{ marginBottom: 12, padding: 12, background: '#fafafa', borderRadius: 6 }}>
          <div style={{ marginBottom: 8 }}>
            <Text strong>查询条件</Text>
            <Text type="secondary" style={{ marginLeft: 8, fontSize: 12 }}>支持模糊查询，多字段同时填写为「且」关系</Text>
          </div>
          <Space wrap>
            {queryFields.map(f => (
              <Input key={f.key} allowClear style={{ width: 220 }} prefix={<SearchOutlined />}
                placeholder={`查询${f.name}`}
                value={fieldQueries[f.key] || ''}
                onChange={e => { setFieldQueries(prev => ({ ...prev, [f.key]: e.target.value })); setPage(1) }} />
            ))}
            <Button onClick={() => { setFieldQueries({}); setPage(1) }}>重置查询</Button>
          </Space>
        </div>
      )}
      <div style={{ marginBottom: 16 }}>
        <Space wrap>
          <Button icon={<ReloadOutlined />} onClick={loadData}>刷新</Button>
          {can('create') && <Button type="primary" icon={<PlusOutlined />} onClick={openCreate}>新增</Button>}
          {can('import') && <Button icon={<UploadOutlined />} onClick={() => { setPreview(null); setImportOpen(true) }}>导入</Button>}
          {can('export') && <Button icon={<DownloadOutlined />} onClick={async () => {
            saveBlob(await api.download(moduleCode, 'export'), `${moduleName}.xlsx`)
          }}>导出</Button>}
          <Button icon={<FileExcelOutlined />} onClick={async () => {
            saveBlob(await api.download(moduleCode, 'template'), `${moduleName}-导入模板.xlsx`)
          }}>下载模板</Button>
        </Space>
        <Text type="secondary" style={{ marginLeft: 12 }}>共 {total} 条 · {fields.length} 个字段（全部展示）</Text>
      </div>

      <Table rowKey="_id" columns={columns} dataSource={rows} loading={loading} size="middle"
        scroll={{ x: 'max-content' }} pagination={{
          current: page, pageSize, total, showSizeChanger: true,
          onChange: (p, ps) => { setPage(p); setPageSize(ps) },
          showTotal: t => `共 ${t} 条`
        }} />

      {/* 新增/编辑弹窗 */}
      <Modal open={editOpen} title={editing ? `编辑 - ${moduleName}` : `新增 - ${moduleName}`}
        onCancel={() => setEditOpen(false)} onOk={submit} width={820}
        okText="保存" cancelText="取消" destroyOnClose>
        <Alert type="info" showIcon style={{ marginBottom: 16 }}
          message="字段依据飞书《各模块字段定义》元数据自动生成。标注「ƒ公式」的字段由系统按口径自动计算、「自动发号」由系统生成、标「自动」的跨表带出，均无需填写。" />
        <DynamicForm fields={fields} initialValues={editing} form={form} />
      </Modal>

      {/* 导入弹窗 */}
      <Modal open={importOpen} title={`导入 - ${moduleName}`} onCancel={() => setImportOpen(false)}
        onOk={commitImport} okText="确认导入有效数据" cancelText="取消" width={900}
        okButtonProps={{ disabled: !preview || preview.valid === 0 }} destroyOnClose>
        <Alert type="info" showIcon style={{ marginBottom: 12 }}
          message="上传 Excel 后先预览校验结果，仅合法行会被导入。校验类型含：枚举值 / 重复性 / 存在性 / 必填 / 数字。" />
        <Upload accept=".xlsx,.xls" beforeUpload={beforeUpload} maxCount={1}
          showUploadList={false}>
          <Button icon={<UploadOutlined />} loading={importing}>选择 Excel 文件</Button>
        </Upload>
        {preview && (
          <div style={{ marginTop: 16 }}>
            <Progress percent={preview.total ? Math.round(preview.valid / preview.total * 100) : 0}
              format={() => `${preview.valid}/${preview.total} 有效`} />
            <Table size="small" rowKey="row" dataSource={preview.rows.slice(0, 50)}
              pagination={{ pageSize: 8 }}
              columns={[
                { title: '行', dataIndex: 'row', width: 50 },
                { title: '结果', dataIndex: 'valid', width: 80, render: v => v ? <Tag color="green">有效</Tag> : <Tag color="red">无效</Tag> },
                { title: '校验信息', key: 'err', render: (_, r) => r.errors.length ? r.errors.join('；') : <Text type="secondary">—</Text> },
              ]} />
          </div>
        )}
      </Modal>
    </div>
  )
}
