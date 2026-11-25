import React, { useEffect, useState } from 'react'
import {
  Layout, Menu, Button, Avatar, Dropdown, Tag, Card, Row, Col, Statistic,
  message, Spin, Empty, Descriptions, Table, Typography, Modal, Space
} from 'antd'
import {
  DatabaseOutlined, FormOutlined, SafetyCertificateOutlined, SettingOutlined,
  LogoutOutlined, UserOutlined, DashboardOutlined, AppstoreOutlined, ReloadOutlined
} from '@ant-design/icons'
import { api, getUsername, setUsername, clearUsername } from './api'
import DataPage from './DataPage'

const { Header, Sider, Content } = Layout
const { Title, Text, Paragraph } = Typography

const TOP_ICON = {
  md: <DatabaseOutlined />, planning: <FormOutlined />,
  calc: <AppstoreOutlined />, perm: <SafetyCertificateOutlined />
}

function Login({ onLogin }) {
  const [username, setU] = useState('admin')
  const [password, setP] = useState('123456')
  const [loading, setL] = useState(false)
  const accounts = [
    ['admin', '管理员', '全部权限'], ['super', '超级管理员', '只查看不改业务'],
    ['budget', '预算管理员', '预算/同步'], ['bp', '财经BP', '填报审核/评审'],
    ['bu', 'BU总', '本BU数据'], ['bg', 'BG总', '本BG数据'], ['cert', '认证人', '认证确认'],
  ]
  const submit = async () => {
    setL(true)
    try {
      const d = await api.login(username, password)
      setUsername(username); onLogin(d)
    } catch (e) { message.error(e.message) }
    finally { setL(false) }
  }
  return (
    <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center',
      background: 'linear-gradient(135deg,#f0f5ff 0%,#e6f7ff 50%,#f6ffed 100%)' }}>
      <Card style={{ width: 460, boxShadow: '0 8px 32px rgba(0,0,0,.12)' }}>
        <div style={{ textAlign: 'center', marginBottom: 24 }}>
          <div style={{ fontSize: 42 }}>📊</div>
          <Title level={3} style={{ margin: '8px 0 4px' }}>产研项目投产规划填报系统</Title>
          <Text type="secondary">主数据 · 投产填报 · 四算 · 权限</Text>
        </div>
        <div style={{ marginBottom: 16 }}>
          <Text strong>快速体验账号（密码均为 123456）：</Text>
          <div style={{ marginTop: 8, display: 'flex', flexWrap: 'wrap', gap: 6 }}>
            {accounts.map(([u, n]) => <Tag key={u} color={u === username ? 'blue' : 'default'}
              style={{ cursor: 'pointer', margin: 0 }} onClick={() => { setU(u); setP('123456') }}>{n}</Tag>)}
          </div>
        </div>
        <div style={{ marginBottom: 12 }}>
          <Text>用户名</Text>
          <input value={username} onChange={e => setU(e.target.value)}
            style={{ width: '100%', padding: '8px 10px', marginTop: 4, border: '1px solid #d9d9d9', borderRadius: 6 }} />
        </div>
        <div style={{ marginBottom: 20 }}>
          <Text>密码</Text>
          <input type="password" value={password} onChange={e => setP(e.target.value)}
            onKeyDown={e => e.key === 'Enter' && submit()}
            style={{ width: '100%', padding: '8px 10px', marginTop: 4, border: '1px solid #d9d9d9', borderRadius: 6 }} />
        </div>
        <Button type="primary" block size="large" loading={loading} onClick={submit}>登录</Button>
      </Card>
    </div>
  )
}

function Dashboard({ meta, goto }) {
  const mdCount = meta.top_modules.find(t => t.code === 'md')?.children.filter(c => c.has_fields).length || 0
  const planCount = meta.top_modules.find(t => t.code === 'planning')?.children.filter(c => c.has_fields).length || 0
  return (
    <div>
      <Title level={3}>工作台</Title>
      <Paragraph type="secondary">欢迎，{meta.user.name}（{meta.user.role_name}）。系统依据飞书《各模块字段定义》构建，当前已落地 {meta.field_count} 个字段、{meta.top_modules.reduce((s, t) => s + t.children.filter(c => c.has_fields).length, 0)} 个模块。</Paragraph>
      <Row gutter={16} style={{ marginBottom: 20 }}>
        <Col span={6}><Card><Statistic title="已落地模块" value={mdCount + planCount} suffix="个" /></Card></Col>
        <Col span={6}><Card><Statistic title="字段总数" value={meta.field_count} suffix="个" /></Card></Col>
        <Col span={6}><Card><Statistic title="主数据模块" value={mdCount} suffix="个" /></Card></Col>
        <Col span={6}><Card><Statistic title="投产填报模块" value={planCount} suffix="个" /></Card></Col>
      </Row>
      <Row gutter={16}>
        <Col span={12}>
          <Card title="模块导航">
            {meta.top_modules.map(top => (
              <div key={top.code} style={{ marginBottom: 12 }}>
                <Text strong>{TOP_ICON[top.code]} {top.name}</Text>
                <div style={{ marginTop: 6, display: 'flex', flexWrap: 'wrap', gap: 8 }}>
                  {top.children.map(c => (
                    <Tag key={c.code} color={c.has_fields ? 'blue' : 'default'}
                      style={{ cursor: 'pointer' }} onClick={() => goto(c)}>
                      {c.name}{!c.has_fields && ' (预留)'}
                    </Tag>
                  ))}
                </div>
              </div>
            ))}
          </Card>
        </Col>
        <Col span={12}>
          <Card title="设计说明">
            <Paragraph><Text strong>元数据驱动：</Text>Sheet2 的 12 列字段定义（类型/必填/查询/导入/导出/枚举/自动取值/校验/格式）固化为注册表，前后端通用组件按此渲染。</Paragraph>
            <Paragraph><Text strong>跨表自动取值：</Text>规划项目等级按 IPD+项目类型从《项目分类标准》带出；生命周期预测的项目类型/IPD 按项目名从《规划项目清单》带出。</Paragraph>
            <Paragraph><Text strong>校验引擎：</Text>枚举值、重复性、存在性、必填、数字格式。</Paragraph>
            <Paragraph><Text strong>权限：</Text>RBAC 功能/按钮 + ABAC 行/字段（切换右上角账号体验）。</Paragraph>
          </Card>
        </Col>
      </Row>
    </div>
  )
}

function ReservedPage({ child }) {
  return (
    <div>
      <Title level={3}>{child.name} <Tag color="orange">预留空位</Tag></Title>
      <Card style={{ marginTop: 16 }}>
        <Empty description={
          <div>
            <Paragraph>该模块在飞书《各模块字段定义》的「功能模块」表中列出，但「二级模块字段」表未提供字段定义。</Paragraph>
            <Paragraph type="secondary">此处保留入口。补充字段定义后即可自动生成 CRUD 页面。</Paragraph>
          </div>
        } />
      </Card>
    </div>
  )
}

function PermissionPage({ meta }) {
  const accounts = [
    ['超级管理员', '系统配置与授权，不直接改业务快照'], ['管理员', '全流程管理与配置'],
    ['预算管理员', '预算属性、同步与对账'], ['财经BP', '填报审核、四算评审'],
    ['BU总', '本 BU 数据查看与审批'], ['BG总', '本 BG 数据查看与审批'], ['认证人', '项目认证与确认'],
  ]
  const cols = [
    { title: '角色', dataIndex: 0, width: 140 },
    { title: '职责', dataIndex: 1 },
  ]
  return (
    <div>
      <Title level={3}>权限管理 <Tag color="orange">演示</Tag></Title>
      <Card title="角色定义（对齐详设 §9）" style={{ marginTop: 16 }}>
        <Table rowKey={0} columns={cols} dataSource={accounts} pagination={false} size="small" />
      </Card>
      <Card title="当前登录上下文" style={{ marginTop: 16 }}>
        <Descriptions bordered column={2} size="small">
          <Descriptions.Item label="姓名">{meta.user.name}</Descriptions.Item>
          <Descriptions.Item label="角色">{meta.user.role_name}</Descriptions.Item>
          <Descriptions.Item label="数据范围(BU)">{meta.user.bu}</Descriptions.Item>
          <Descriptions.Item label="数据范围(BG)">{meta.user.bg}</Descriptions.Item>
          <Descriptions.Item label="可用按钮" span={2}>
            {meta.user.buttons.map(b => <Tag key={b} color="green">{b}</Tag>)}
          </Descriptions.Item>
        </Descriptions>
        <Paragraph type="secondary" style={{ marginTop: 12 }}>
          切换右上角「退出登录」后用不同账号登录，可体验功能权限（可见模块）、按钮权限（增改导入导出）、行权限（BU/BG 数据范围）与字段权限（税率等敏感字段脱敏）。
        </Paragraph>
      </Card>
    </div>
  )
}

export default function App() {
  const [user, setUser] = useState(null)
  const [meta, setMeta] = useState(null)
  const [current, setCurrent] = useState(null)
  const [booting, setBooting] = useState(true)

  const bootstrap = async () => {
    if (!getUsername()) { setBooting(false); return }
    try { setMeta(await api.meta()); setUser({ name: getUsername() }) }
    catch { clearUsername() }
    finally { setBooting(false) }
  }
  useEffect(() => { bootstrap() }, [])

  const onLogin = async () => {
    setMeta(await api.meta()); setUser({ name: getUsername() }); setCurrent(null)
  }
  const onLogout = () => { clearUsername(); setUser(null); setMeta(null); setCurrent(null) }
  const goto = (child) => setCurrent(child)

  if (booting) return <div style={{ height: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center' }}><Spin size="large" /></div>
  if (!user || !meta) return <Login onLogin={onLogin} />

  const menuItems = meta.top_modules.map(top => ({
    key: `top_${top.code}`,
    icon: TOP_ICON[top.code],
    label: top.name,
    children: top.children.map(c => ({
      key: c.code,
      label: <span>{c.name}{!c.has_fields && <Tag color="orange" style={{ marginLeft: 6, fontSize: 10 }}>预留</Tag>}</span>,
    }))
  }))
  menuItems.push({ key: '_dashboard', icon: <DashboardOutlined />, label: '工作台' })
  menuItems.push({ key: '_perm', icon: <SafetyCertificateOutlined />, label: '权限管理' })

  const onMenu = ({ key }) => {
    if (key === '_dashboard') { setCurrent(null); return }
    if (key === '_perm') { setCurrent({ code: '_perm', name: '权限管理' }); return }
    if (key.startsWith('top_')) return
    let found = null
    meta.top_modules.forEach(t => t.children.forEach(c => { if (c.code === key) found = c }))
    if (found) goto(found)
  }

  const renderContent = () => {
    if (!current) return <Dashboard meta={meta} goto={goto} />
    if (current.code === '_perm') return <PermissionPage meta={meta} />
    if (!current.has_fields) return <ReservedPage child={current} />
    return <DataPage moduleCode={current.code} moduleName={current.name} user={meta.user} />
  }

  return (
    <Layout style={{ minHeight: '100vh' }}>
      <Sider width={240} theme="light" style={{ borderRight: '1px solid #f0f0f0' }}>
        <div style={{ height: 60, display: 'flex', alignItems: 'center', paddingLeft: 20, borderBottom: '1px solid #f0f0f0' }}>
          <span style={{ fontSize: 24, marginRight: 8 }}>📊</span>
          <Text strong style={{ fontSize: 15 }}>投产规划填报系统</Text>
        </div>
        <Menu mode="inline" items={menuItems} onClick={onMenu}
          selectedKeys={current ? [current.code] : ['_dashboard']}
          defaultOpenKeys={['top_md', 'top_planning']} />
      </Sider>
      <Layout>
        <Header style={{ background: '#fff', padding: '0 24px', display: 'flex', alignItems: 'center', justifyContent: 'flex-end', borderBottom: '1px solid #f0f0f0' }}>
          <Space>
            <Tag color="blue">{meta.user.role_name}</Tag>
            <Dropdown menu={{ items: [
              { key: 'reset', label: '重置模拟数据', icon: <ReloadOutlined /> },
              { type: 'divider' },
              { key: 'logout', label: '退出登录', icon: <LogoutOutlined /> },
            ], onClick: async ({ key }) => {
              if (key === 'logout') onLogout()
              if (key === 'reset') {
                Modal.confirm({
                  title: '重置模拟数据', content: '将清空并重新生成所有模块的演示数据，确认继续？',
                  onOk: async () => { await api.resetSeed(); message.success('已重置') }
                })
              }
            } }}>
              <Space style={{ cursor: 'pointer' }}>
                <Avatar icon={<UserOutlined />} /> {meta.user.name}
              </Space>
            </Dropdown>
          </Space>
        </Header>
        <Content style={{ margin: 20, padding: 24, background: '#fff', borderRadius: 8, minHeight: 360 }}>
          {renderContent()}
        </Content>
      </Layout>
    </Layout>
  )
}
