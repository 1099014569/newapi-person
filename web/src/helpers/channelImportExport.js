import {
  API,
  showError,
  showSuccess,
  showInfo,
  getUserIdFromLocalStorage,
} from './index';

/**
 * 导出渠道到 Excel 文件
 */
export const exportChannels = async () => {
  try {
    const response = await fetch('/api/channel/export', {
      method: 'GET',
      headers: {
        Authorization: localStorage.getItem('token'),
        'New-API-User': getUserIdFromLocalStorage().toString(),
      },
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.message || '导出失败');
    }

    // 获取文件名
    const contentDisposition = response.headers.get('Content-Disposition');
    let filename = 'channels_export.xlsx';
    if (contentDisposition) {
      const filenameMatch = contentDisposition.match(/filename=([^;]+)/);
      if (filenameMatch) {
        filename = filenameMatch[1].trim();
      }
    }

    // 下载文件
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);

    showSuccess('导出成功');
    return true;
  } catch (error) {
    showError('导出失败: ' + error.message);
    return false;
  }
};

/**
 * 导入渠道从 Excel 文件
 * @param {File} file - Excel 文件
 * @param {Function} onSuccess - 成功回调
 */
export const importChannels = async (file, onSuccess) => {
  try {
    const formData = new FormData();
    formData.append('file', file);

    const response = await API.post('/api/channel/import', formData, {
      headers: {
        'Content-Type': 'multipart/form-data',
      },
    });

    const { success, message, data } = response.data;

    if (success) {
      const result = {
        total: data.total || 0,
        successCount: data.success_count || 0,
        failCount: data.fail_count || 0,
        errors: data.errors || [],
      };

      if (result.failCount > 0) {
        showInfo(message || '导入完成，部分记录失败');
      } else {
        showSuccess(message || '导入成功');
      }

      if (onSuccess) {
        onSuccess(result);
      }

      return result;
    } else {
      showError(message || '导入失败');
      return null;
    }
  } catch (error) {
    showError('导入失败: ' + (error.response?.data?.message || error.message));
    return null;
  }
};
