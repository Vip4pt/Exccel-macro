function 按列进行单元格合并() {
    try {
        let sheet = Application.ActiveSheet;
        let startRow = 380; // 起始行
        let mergeInterval = 6; // 合并间隔
        let endRow = 1000; // 请手动指定一个预期的结束行号
        let columnsToMerge = [1, 2, 3]; // 列: A=1, B=2, C=3

        // 关闭屏幕更新以提高性能
        Application.ScreenUpdating = false;

        let mergedCount = 0;
        
        // 循环处理每一行和每一列
        for (let row = startRow; row <= endRow; row += mergeInterval) {
            // 检查是否超出工作表实际行数
            if (row + mergeInterval - 1 > sheet.Rows.Count) {
                Debug.Print("超出工作表行数，停止在行: " + row);
                break;
            }
            
            for (let colIndex of columnsToMerge) {
                // 定义要合并的单元格范围
                let mergeRange = sheet.Range(
                    sheet.Cells(row, colIndex),
                    sheet.Cells(row + mergeInterval - 1, colIndex)
                );
                
                // 先判断一下这个区域是否已经被合并了
                if (!mergeRange.MergeCells) {
                    mergeRange.Merge();
                    mergedCount++;
                    Debug.Print("合并范围: 行 " + row + " 到 " + (row + mergeInterval - 1) + ", 列 " + colIndex);
                } else {
                    Debug.Print("跳过已合并区域: 行 " + row + " 到 " + (row + mergeInterval - 1) + ", 列 " + colIndex);
                }
            }
        }
        
        Debug.Print("自动合并完成！共处理了 " + mergedCount + " 个合并区域。");
        
    } catch (error) {
        Debug.Print("执行过程中发生错误: " + error.message);
    } finally {
        // 恢复屏幕更新
        Application.ScreenUpdating = true;
    }
}
