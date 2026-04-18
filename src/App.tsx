import { useState, useEffect } from 'react'
import { bitable, IFieldMeta } from '@lark-base-open/js-sdk'
import './App.css'

function App() {
  const [tableName, setTableName] = useState<string>('Loading...')
  const [recordCount, setRecordCount] = useState<number>(0)
  const [isExporting, setIsExporting] = useState<boolean>(false)
  const [statusMsg, setStatusMsg] = useState<string>('')
  const [logs, setLogs] = useState<{ msg: string; type: 'info' | 'success' | 'error' }[]>([])
  
  const [fields, setFields] = useState<IFieldMeta[]>([])
  const [selectedFieldIds, setSelectedFieldIds] = useState<string[]>([])
  const [exportFormat, setExportFormat] = useState<'csv' | 'json'>('csv')

  useEffect(() => {
    const fetchData = async () => {
      try {
        const table = await bitable.base.getActiveTable()
        const name = await table.getName()
        setTableName(name)

        const activeView = await table.getActiveView()
        const viewFieldMetaList = await activeView.getFieldMetaList()
        setFields(viewFieldMetaList)
        setSelectedFieldIds(viewFieldMetaList.map(f => f.id))
        
        const visibleRecordIds = await activeView.getVisibleRecordIdList()
        setRecordCount(visibleRecordIds.filter(id => id).length)
      } catch (error) {
        console.error('Failed to fetch data:', error)
        try {
          const table = await bitable.base.getActiveTable()
          const allFields = await table.getFieldMetaList()
          setFields(allFields)
          setSelectedFieldIds(allFields.map(f => f.id))
          const recordList = await table.getRecordIdList()
          setRecordCount(recordList.length)
        } catch (e) {
          console.error('Fallback also failed:', e)
        }
      }
    }

    fetchData()

    const off = bitable.base.onSelectionChange(async (event) => {
      if (event.data.tableId) {
        fetchData()
      }
    })

    return () => off()
  }, [])

  const addLog = (msg: string, type: 'info' | 'success' | 'error' = 'info') => {
    setLogs(prev => [{ msg, type }, ...prev].slice(0, 50))
  }

  const toggleFieldSelection = (fieldId: string) => {
    setSelectedFieldIds(prev => {
      if (prev.includes(fieldId)) {
        return prev.filter(id => id !== fieldId)
      } else {
        return [...prev, fieldId]
      }
    })
  }

  const selectAllFields = () => {
    if (selectedFieldIds.length === fields.length) {
      setSelectedFieldIds([])
    } else {
      setSelectedFieldIds(fields.map(f => f.id))
    }
  }

  const escapeCSV = (value: any): string => {
    if (value === null || value === undefined) return ''
    const str = String(value)
    if (str.includes(',') || str.includes('"') || str.includes('\n')) {
      return '"' + str.replace(/"/g, '""') + '"'
    }
    return str
  }

  const getCellTextValue = async (fieldId: string, recordId: string): Promise<string> => {
    try {
      const table = await bitable.base.getActiveTable()
      const field = await table.getField(fieldId)
      const value = await field.getValue(recordId)
      const fieldMeta = fields.find(f => f.id === fieldId)
      const fieldType = fieldMeta?.type || 0
      return formatCellValue(fieldType, value)
    } catch (error) {
      console.error('Error getting cell value:', error)
      return ''
    }
  }

  const formatCellValue = (fieldType: number, value: any): string => {
    if (value === null || value === undefined) return ''
    
    if (typeof value === 'string') return value
    if (typeof value === 'number') return String(value)
    if (typeof value === 'boolean') return value ? '是' : '否'
    
    if (Array.isArray(value)) {
      if (value.length === 0) return ''
      const items = value.map((item: any) => {
        if (typeof item === 'string') return item
        if (typeof item === 'object' && item !== null) {
          return item.name || item.text || item.value || JSON.stringify(item)
        }
        return String(item)
      })
      return items.join(';')
    }
    
    if (typeof value === 'object') {
      return value.name || value.text || value.value || JSON.stringify(value)
    }
    
    return String(value)
  }

  const exportToCSV = async (): Promise<string> => {
    const table = await bitable.base.getActiveTable()
    const activeView = await table.getActiveView()
    
    const selectedFields = fields.filter(f => selectedFieldIds.includes(f.id))
    const headers = selectedFields.map(f => f.name)
    
    const visibleRecordIds = await activeView.getVisibleRecordIdList()
    const recordIds = visibleRecordIds.filter(id => id) as string[]

    const fieldObjects = await Promise.all(
      selectedFields.map(async (field) => ({
        id: field.id,
        field: await table.getField(field.id),
        type: field.type
      }))
    )

    const rows: string[] = [headers.map(escapeCSV).join(',')]

    const batchSize = 100
    for (let i = 0; i < recordIds.length; i += batchSize) {
      const batchRecordIds = recordIds.slice(i, i + batchSize)
      const batchPromises = batchRecordIds.map(async (recordId) => {
        const values = await Promise.all(
          fieldObjects.map(async (f) => {
            try {
              const value = await f.field.getValue(recordId)
              return formatCellValue(f.type, value)
            } catch {
              return ''
            }
          })
        )
        return values.map(escapeCSV).join(',')
      })
      const batchResults = await Promise.all(batchPromises)
      rows.push(...batchResults)
      addLog(`已处理 ${Math.min(i + batchSize, recordIds.length)}/${recordIds.length} 条记录`, 'info')
    }
    
    return '\uFEFF' + rows.join('\n')
  }

  const exportToJSON = async (): Promise<string> => {
    const table = await bitable.base.getActiveTable()
    const activeView = await table.getActiveView()
    
    const selectedFields = fields.filter(f => selectedFieldIds.includes(f.id))
    
    const visibleRecordIds = await activeView.getVisibleRecordIdList()
    const recordIds = visibleRecordIds.filter(id => id) as string[]

    const fieldObjects = await Promise.all(
      selectedFields.map(async (field) => ({
        id: field.id,
        name: field.name,
        field: await table.getField(field.id),
        type: field.type
      }))
    )

    const data: any[] = []

    const batchSize = 100
    for (let i = 0; i < recordIds.length; i += batchSize) {
      const batchRecordIds = recordIds.slice(i, i + batchSize)
      const batchPromises = batchRecordIds.map(async (recordId) => {
        const row: Record<string, any> = { _recordId: recordId }
        const values = await Promise.all(
          fieldObjects.map(async (f) => {
            try {
              const value = await f.field.getValue(recordId)
              return [f.name, formatCellValue(f.type, value)]
            } catch {
              return [f.name, '']
            }
          })
        )
        values.forEach(([name, val]) => {
          row[name as string] = val
        })
        return row
      })
      const batchResults = await Promise.all(batchPromises)
      data.push(...batchResults)
      addLog(`已处理 ${Math.min(i + batchSize, recordIds.length)}/${recordIds.length} 条记录`, 'info')
    }
    
    return JSON.stringify(data, null, 2)
  }

  const handleExport = async () => {
    if (selectedFieldIds.length === 0) {
      setStatusMsg('请至少选择一个字段')
      return
    }

    setIsExporting(true)
    setStatusMsg('正在导出...')
    setLogs([])

    try {
      let content: string
      let fileName: string
      let mimeType: string

      if (exportFormat === 'csv') {
        content = await exportToCSV()
        fileName = `${tableName}_${new Date().toISOString().slice(0, 10)}.csv`
        mimeType = 'text/csv;charset=utf-8'
      } else {
        content = await exportToJSON()
        fileName = `${tableName}_${new Date().toISOString().slice(0, 10)}.json`
        mimeType = 'application/json'
      }

      const blob = new Blob([content], { type: mimeType })
      const url = URL.createObjectURL(blob)
      const link = document.createElement('a')
      link.href = url
      link.download = fileName
      link.click()
      URL.revokeObjectURL(url)

      addLog(`导出成功: ${fileName}`, 'success')
      setStatusMsg(`导出成功！已导出 ${recordCount} 条记录，${selectedFieldIds.length} 个字段`)
    } catch (error: any) {
      console.error(error)
      setStatusMsg(`导出失败: ${error.message || '未知错误'}`)
      addLog(`错误: ${error.message}`, 'error')
    } finally {
      setIsExporting(false)
    }
  }

  return (
    <div className="container">
      <h1>数据导出</h1>
      
      <div className="card">
        <h3>📊 表格信息</h3>
        <p>当前表: <strong>{tableName}</strong></p>
        <p>记录数: <strong>{recordCount}</strong></p>
        <p>字段数: <strong>{fields.length}</strong></p>
      </div>

      <div className="card">
        <h3>📥 导出数据</h3>
        <p className="desc">选择字段，导出为 CSV 或 JSON 格式</p>
        
        <div className="form-group">
          <label>导出格式</label>
          <div className="radio-group">
            <label>
              <input 
                type="radio" 
                name="format" 
                checked={exportFormat === 'csv'}
                onChange={() => setExportFormat('csv')}
              />
              CSV (Excel)
            </label>
            <label>
              <input 
                type="radio" 
                name="format" 
                checked={exportFormat === 'json'}
                onChange={() => setExportFormat('json')}
              />
              JSON
            </label>
          </div>
        </div>

        <div className="record-list-header">
          <label className="checkbox-label">
            <input 
              type="checkbox" 
              checked={selectedFieldIds.length === fields.length && fields.length > 0}
              onChange={selectAllFields}
            />
            <span>全选字段 ({selectedFieldIds.length}/{fields.length})</span>
          </label>
        </div>

        <div className="record-list">
          {fields.length > 0 ? (
            fields.map((field) => (
              <div 
                key={field.id} 
                className={`record-item ${selectedFieldIds.includes(field.id) ? 'selected' : ''}`}
                onClick={() => toggleFieldSelection(field.id)}
              >
                <input 
                  type="checkbox" 
                  checked={selectedFieldIds.includes(field.id)}
                  onChange={() => {}}
                />
                <span>{field.name}</span>
              </div>
            ))
          ) : (
            <p className="no-records">暂无字段</p>
          )}
        </div>

        <div className="button-group">
          <button 
            onClick={handleExport} 
            disabled={isExporting || selectedFieldIds.length === 0 || fields.length === 0}
            className={`convert-btn ${isExporting ? 'loading' : ''}`}
          >
            {isExporting ? '导出中...' : `导出 ${selectedFieldIds.length} 个字段`}
          </button>
        </div>
        {statusMsg && <p className={`status-msg ${statusMsg.includes('成功') ? 'success' : 'error'}`}>{statusMsg}</p>}

        {logs.length > 0 && (
          <div className="log-container">
            <h4>日志</h4>
            <div className="log-list">
              {logs.map((log, index) => (
                <div key={index} className={`log-item ${log.type}`}>
                  {log.msg}
                </div>
              ))}
            </div>
          </div>
        )}
      </div>

      <p className="footer">
        山东代理区-数据运营开发<br/>
        有问题联系裴忠瀚
      </p>
    </div>
  )
}

export default App
