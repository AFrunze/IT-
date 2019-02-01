--Установим драйвер Microsoft.ACE.OLEDB.12.0
--необходимо включить параметр ad hoc distributed queries в конфигурации сервера
sp_configure 'show advanced options', 1;
RECONFIGURE;
GO
sp_configure 'Ad Hoc Distributed Queries', 1;
RECONFIGURE;
GO

--В приведенном ниже примере кода данные импортируются из листа Excel base в новую таблицу базы данных с помощью OPENROWSET
USE ImportFromExcel;
GO
SELECT * INTO base_dq
FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
    'Excel 12.0; Database=D:\123\1.xlsx', [base$]);
GO
