DECLARE @path VARCHAR(50);
DECLARE @dir VARCHAR(50);
DECLARE @shell VARCHAR(50);
DECLARE @restoreToDataDir nvarchar(500)

-- where you want to save .mdf and log files
SET @restoreToDataDir = 'c:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA\';
-- path to directory
SET @path = 'C:\SQLBackups\';
SET @dir = 'dir /B ';
SET @shell = @dir + @path;
CREATE TABLE tmp(excelFileName VARCHAR(100));
INSERT INTO tmp
EXEC xp_cmdshell @shell

declare @fileName varchar(100)
declare @sfileName varchar(100)

While (Select Count(*) From tmp where excelFileName is not null) > 0
Begin

    Select Top 1 @fileName = excelFileName From tmp
   	-- remove last 13 chars..
   	SET @sfileName =  LEFT(@fileName, LEN(@fileName) - 13) 
	--PRINT(@sfileName)
	
	DECLARE @fileListTable table
(
    LogicalName          nvarchar(128),
    PhysicalName         nvarchar(260),
    [Type]               char(1),
    FileGroupName        nvarchar(128),
    Size                 numeric(20,0),
    MaxSize              numeric(20,0),
    FileID               bigint,
    CreateLSN            numeric(25,0),
    DropLSN              numeric(25,0),
    UniqueID             uniqueidentifier,
    ReadOnlyLSN          numeric(25,0),
    ReadWriteLSN         numeric(25,0),
    BackupSizeInBytes    bigint,
    SourceBlockSize      int,
    FileGroupID          int,
    LogGroupGUID         uniqueidentifier,
    DifferentialBaseLSN  numeric(25,0),
    DifferentialBaseGUID uniqueidentifier,
    IsReadOnl            bit,
    IsPresent            bit,
    TDEThumbprint        varbinary(32) -- remove this column if using SQL 2005,

	)
	DECLARE @exec VARCHAR(500);
	DECLARE @sexec VARCHAR(500);
	SET @exec = 'RESTORE FILELISTONLY FROM disk = ''' + @path +  @fileName + '''' 
	PRINT(@exec)

	INSERT into @fileListTable exec(@exec)

	SELECT * from @fileListTable
		-- OPENROWSET processing goes here, using @fileName to identify which file to use
		DECLARE @restore VARCHAR(500)
		DECLARE @LogicalName VARCHAR(500)
		DECLARE @LogicalName_log NVARCHAR(500)
		DECLARE @PhysicalName NVARCHAR(500)
		DECLARE @PhysicalName_log NVARCHAR(500)
	    
		SET @restore = @path + @fileName
		SELECT  TOP 1 @LogicalName = LogicalName  from @fileListTable
		select TOP 1 @LogicalName_log = LogicalName from @fileListTable ORDER BY LogicalName DESC;
	    
		SET @PhysicalName = @restoreToDataDir+@LogicalName+'.MDF'
		SET @PhysicalName_log = @restoreToDataDir+@LogicalName_log+'.LDF'
		--PRINT(@PhysicalName)
	    
	RESTORE DATABASE @sfileName
	FROM DISK = @restore
	WITH MOVE @LogicalName TO @PhysicalName,
	MOVE @LogicalName_log TO @PhysicalName_log



	DELETE from tmp Where excelFileName = @FileName
	DELETE FROM @fileListTable

End

DROP TABLE tmp