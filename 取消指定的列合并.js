function 取消指定合并() {
    let sheet = Application.ActiveSheet;
    let startRow = 386; // 起始行
    let endRow = 1000; // 结束行
    let columnsToUnmerge = [2]; // 要取消合并的列: A=1, B=2, C=3
    
    // 关闭屏幕更新以提高性能
    Application.ScreenUpdating = false;
    
    try {
        // 循环处理每一列
        for (let colIndex of columnsToUnmerge) {
            // 循环处理每一行
            for (let row = startRow; row <= endRow; row++) {
                let cell = sheet.Cells(row, colIndex);
                
                // 检查单元格是否被合并
                if (cell.MergeCells) {
                    // 获取合并区域
                    let mergeArea = cell.MergeArea;
                    
                    // 取消合并
                    mergeArea.UnMerge();
                    
                    // 可选: 清除合并区域的内容
                    // mergeArea.ClearContents();
                }
            }
        }
    } catch (error) {
        // 错误处理
    } finally {
        // 恢复屏幕更新
        Application.ScreenUpdating = true;
    }
}
