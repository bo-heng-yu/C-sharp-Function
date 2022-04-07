/// <summary>
/// 讀取excel樣板，將撈出的DB DataTable寫入excel 儲存在記憶體 傳出
/// Author : Heng
/// Update : Tone
/// Create Time : 2021/07/07/11:00 
/// </summary>
/// <param name="excelItem"></param> excel資訊
/// <param name="sqlResult"></param> DB撈出的table
/// <returns></returns>
public byte[] GetExcelContent(DataTable sqlResult, Dictionary<string,object> excelItem)
{
    try{
        byte[] content;
        var envParams       = GetEnvParameterList();
        var sExcelFolder    = envParams["sampleExcelFolder"].ToString();

        using (var workbook = new XLWorkbook($"{this._mainPath}{sExcelFolder}\\{excelItem["excelName"]}.xlsx"))
        {
            var worksheet = workbook.Worksheet(1); //選excel 第一張 table
            foreach(var excelData in excelItem){
                var excelCol = excelData.Key;
                var excelVal = excelData.Value;
                var findCells =  worksheet.Search("#"+excelCol);
                if(findCells.Count() > 0){
                    var originText  = findCells.First().GetValue<string>();
                    var newText     = originText.Replace("#"+excelCol, excelVal == null ? "" : excelVal.ToString());
                    findCells.First().SetValue(newText);
                }
            }
            if(worksheet.RowsUsed().Count() > 0){
                var startRow = worksheet.LastRowUsed().RowNumber() + 1;  //已使用過的行Row,從下個空白列開始,所以+1
                worksheet.Cell($"A{startRow}").InsertTable(sqlResult); //將DB的table寫入
            }else{
                worksheet.Cell("A1").InsertTable(sqlResult); //將DB的table寫入
            }

            using (var stream = new MemoryStream()){ //存入記憶體 寫入content傳出
                workbook.SaveAs(stream);
                content = stream.ToArray();
                return content;
            }
        }
    }catch(Exception e){
        Console.WriteLine($"error:{e.ToString()}");
        throw e;
    }
}
/// <summary>
/// 讀取excel樣板，將撈出的DB 多筆DataTable個別寫入excel table 儲存在記憶體 傳出
/// Author : Heng
/// Create Time : 2021/11/25/17:00 
/// </summary>
/// <param name="excelItem"></param> excel資訊
/// <param name="sqlResult"></param> DB撈出的dataSet
/// <returns></returns>
public byte[] GetExcelContent(DataSet sqlResult, Dictionary<string,object> excelItem)
{
    try{
        byte[] content;
        var envParams       = GetEnvParameterList();
        var sExcelFolder    = envParams["sampleExcelFolder"].ToString();
        using (var workbook = new XLWorkbook($"{this._mainPath}{sExcelFolder}\\{excelItem["excelName"]}.xlsx"))
        {
            int index = 1;
            foreach(DataTable dataTable in sqlResult.Tables){
                if(index > 1)
                    workbook.AddWorksheet(index);           //先取名再新增table
                var worksheet = workbook.Worksheet(index);  //選excel table位置
                foreach(var excelData in excelItem){
                    var excelCol = excelData.Key;
                    var excelVal = excelData.Value;
                    var findCells =  worksheet.Search("#"+excelCol);
                    if(findCells.Count() > 0){
                        var originText  = findCells.First().GetValue<string>();
                        var newText     = originText.Replace("#"+excelCol, excelVal == null ? "" : excelVal.ToString());
                        findCells.First().SetValue(newText);
                    }
                }
                if(worksheet.RowsUsed().Count() > 0){
                    var startRow = worksheet.LastRowUsed().RowNumber() + 1;  //已使用過的行Row,從下個空白列開始,所以+1
                    worksheet.Cell($"A{startRow}").InsertTable(dataTable); //將DB的table寫入
                }else{
                    worksheet.Cell("A1").InsertTable(dataTable); //將DB的table寫入
                }
                index ++;
            }
            using (var stream = new MemoryStream()){ //存入記憶體 寫入content傳出
                workbook.SaveAs(stream);
                content = stream.ToArray();
                return content;
            }
        }
    }catch(Exception e){
        Console.WriteLine($"error:{e.ToString()}");
        throw e;
    }
}
