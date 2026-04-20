/**
 * 校验文件格式是否支持（PDF/Word/Excel/TXT/Markdown）
 * @param {File} file - 待校验文件
 * @returns {boolean} - 是否支持
 */
export function checkFileFormat(file) {
    const supportTypes = [
        'application/pdf',
        'application/msword',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'application/vnd.ms-excel',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'text/plain',
        'text/markdown'
    ]
    return supportTypes.includes(file.type)
}

/**
 * 格式化文件大小（字节转MB）
 * @param {number} size - 文件大小（字节）
 * @returns {string} - 格式化后的大小（如：2.5MB）
 */
export function formatFileSize(size) {
    return (size / 1024 / 1024).toFixed(2) + 'MB'
}