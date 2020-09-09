
DROP TABLE #該修改的欄位名稱
DROP TABLE #學校代碼
DROP TABLE #temp1
	DROP TABLE #上傳紀錄數據


SELECT  DISTINCT [學校代碼] INTO #temp1 FROM [歷程上傳].[dbo].[_108入學學生上傳]
SELECT ROW_NUMBER() OVER(ORDER BY [學校代碼] ASC) AS Row, [學校代碼] INTO #學校代碼 FROM #temp1
DROP TABLE #temp1
DECLARE @學校數量 AS INTEGER = (SELECT MAX(Row) FROM #學校代碼)
DECLARE @欄位數量 AS INTEGER
DECLARE @紀錄數量 AS INTEGER

DECLARE @i_學校索引數 AS INTEGER, @j_單個學校紀錄索引數 AS INTEGER, @k_欄位索引數 AS INTEGER
DECLARE @current AS INTEGER
DECLARE @previous AS INTEGER
DECLARE @previousValue AS float
DECLARE @學校代碼 AS VARCHAR(max)
DECLARE @欄位名稱 AS VARCHAR(max)
DECLARE @value AS FLOAT
DECLARE @cmd AS VARCHAR(max)
USE 歷程上傳
SELECT ROW_NUMBER() OVER(ORDER BY COLUMN_NAME ASC) AS Row, COLUMN_NAME, DATA_TYPE INTO #該修改的欄位名稱
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = '_108入學學生上傳' AND TABLE_SCHEMA='dbo' AND DATA_TYPE IN ('int', 'float') AND COLUMN_NAME NOT IN('編號', 'sid')
SET @欄位數量 = (SELECT MAX(Row) FROM #該修改的欄位名稱)
--SELECT * FROM #該修改的欄位名稱
--SELECT @欄位數量

SET @i_學校索引數 = 1
WHILE @i_學校索引數 <= @學校數量
BEGIN
	SET @學校代碼 = (SELECT [學校代碼] FROM #學校代碼 WHERE Row = @i_學校索引數)
	SELECT ROW_NUMBER() OVER(ORDER BY [統計日期] ASC) AS Row, * INTO #上傳紀錄數據 FROM [歷程上傳].[dbo].[_108入學學生上傳]  WHERE [學校代碼] = @學校代碼
	--SELECT * FROM #上傳紀錄數據
	SET @j_單個學校紀錄索引數 = 2
	SET @紀錄數量 = (SELECT MAX(Row) FROM #上傳紀錄數據)
	WHILE @j_單個學校紀錄索引數 <= @紀錄數量
	BEGIN
		SET @current = (SELECT [sid] FROM #上傳紀錄數據 WHERE Row = @j_單個學校紀錄索引數)
		SET @previous = (SELECT [sid] FROM #上傳紀錄數據 WHERE Row = (@j_單個學校紀錄索引數 - 1))
		--SELECT @current
		--SELECT @current, @previous
		SET @k_欄位索引數 = 1
		WHILE @k_欄位索引數 <= @欄位數量
		BEGIN
			SET @欄位名稱 = (SELECT COLUMN_NAME FROM #該修改的欄位名稱 WHERE Row = @k_欄位索引數)

			CREATE TABLE #result ([rowcount] FLOAT);
			INSERT INTO #result ([rowcount])
			EXEC (N'SELECT ' + @欄位名稱 + '  FROM #上傳紀錄數據 WHERE Row = ' + @j_單個學校紀錄索引數);
			SET @value = (select top (1) [rowcount] from #result);
			DROP TABLE #result

			IF @value = 0 OR @value IS NULL
			BEGIN
				SET @cmd = 'UPDATE ' +
								'[_108入學學生上傳] ' +
							'SET ' + @欄位名稱 +
							' = (SELECT ' + @欄位名稱 + ' FROM [_108入學學生上傳] WHERE sid = ' + CAST(@previous AS VARCHAR) + ')' +
							'WHERE sid = ' + CAST(@current AS VARCHAR);
				EXEC(@cmd)
			END

			SET @k_欄位索引數 += 1
		END
		SET @j_單個學校紀錄索引數 += 1
	END
	SET @i_學校索引數 += 1
	DROP TABLE #上傳紀錄數據
END
DROP TABLE #該修改的欄位名稱
DROP TABLE #學校代碼
