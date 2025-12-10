/*
Copyright (C) 2025 QuantumNous

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as
published by the Free Software Foundation, either version 3 of the
License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License
along with this program. If not, see <https://www.gnu.org/licenses/>.

For commercial licensing, please contact support@quantumnous.com
*/

import React from 'react';
import { Modal, Table, Tag } from '@douyinfe/semi-ui';

const ImportResultModal = ({ visible, result, onClose, t }) => {
  if (!result) return null;

  const columns = [
    {
      title: t('行号'),
      dataIndex: 'row',
      key: 'row',
      width: 80,
    },
    {
      title: t('渠道名称'),
      dataIndex: 'name',
      key: 'name',
      width: 150,
    },
    {
      title: t('错误信息'),
      dataIndex: 'message',
      key: 'message',
      render: (text) => <span style={{ color: '#f53f3f' }}>{text}</span>,
    },
  ];

  return (
    <Modal
      title={t('导入结果')}
      visible={visible}
      onCancel={onClose}
      onOk={onClose}
      cancelText={null}
      okText={t('关闭')}
      width={800}
      bodyStyle={{ padding: '20px' }}
    >
      <div style={{ marginBottom: 20 }}>
        <div style={{ marginBottom: 10 }}>
          <Tag color='blue' size='large'>
            {t('总计')}：{result.total}
          </Tag>
          <Tag color='green' size='large' style={{ marginLeft: 10 }}>
            {t('成功')}：{result.successCount}
          </Tag>
          {result.failCount > 0 && (
            <Tag color='red' size='large' style={{ marginLeft: 10 }}>
              {t('失败')}：{result.failCount}
            </Tag>
          )}
        </div>
      </div>

      {result.errors && result.errors.length > 0 && (
        <>
          <h4 style={{ marginTop: 20, marginBottom: 10 }}>
            {t('失败详情')}：
          </h4>
          <Table
            columns={columns}
            dataSource={result.errors}
            pagination={false}
            size='small'
            style={{ maxHeight: 400, overflow: 'auto' }}
          />
        </>
      )}

      {result.errors && result.errors.length === 0 && (
        <div
          style={{
            textAlign: 'center',
            padding: 20,
            color: '#52c41a',
            fontSize: 16,
          }}
        >
          ✓ {t('所有渠道导入成功')}！
        </div>
      )}
    </Modal>
  );
};

export default ImportResultModal;
