-- =============================================
-- SaveToDB Administrator Framework for Microsoft SQL Server
-- Version 10.6, December 13, 2022
--
-- Copyright 2018-2022 Gartle LLC
--
-- License: MIT
-- =============================================

SET NOCOUNT ON

DELETE FROM xls.formats   WHERE TABLE_SCHEMA = 'xls' AND TABLE_NAME IN ('usp_database_permissions', 'usp_object_permissions', 'usp_principal_permissions', 'usp_role_members');
DELETE FROM xls.handlers  WHERE TABLE_SCHEMA = 'xls' AND TABLE_NAME IN ('usp_database_permissions', 'usp_object_permissions', 'usp_principal_permissions', 'usp_role_members', 'administrator_framework');
DELETE FROM xls.workbooks WHERE TABLE_SCHEMA = 'xls' AND NAME IN ('savetodb_permissions.xlsx');
GO

DECLARE @id int

SET @id = COALESCE((SELECT MAX(ID) FROM xls.formats), 0);

DBCC CHECKIDENT ('xls.formats', RESEED, @id) WITH NO_INFOMSGS;

SET @id = COALESCE((SELECT MAX(ID) FROM xls.handlers), 0);

DBCC CHECKIDENT ('xls.handlers', RESEED, @id) WITH NO_INFOMSGS;

SET @id = COALESCE((SELECT MAX(ID) FROM xls.workbooks), 0);

DBCC CHECKIDENT ('xls.workbooks', RESEED, @id) WITH NO_INFOMSGS;
GO


IF OBJECT_ID('[xls].[usp_database_permissions]', 'P') IS NOT NULL
DROP PROCEDURE [xls].[usp_database_permissions];
GO
IF OBJECT_ID('[xls].[usp_database_permissions_change]', 'P') IS NOT NULL
DROP PROCEDURE [xls].[usp_database_permissions_change];
GO
IF OBJECT_ID('[xls].[usp_object_permissions]', 'P') IS NOT NULL
DROP PROCEDURE [xls].[usp_object_permissions];
GO
IF OBJECT_ID('[xls].[usp_object_permissions_change]', 'P') IS NOT NULL
DROP PROCEDURE [xls].[usp_object_permissions_change];
GO
IF OBJECT_ID('[xls].[usp_principal_permissions]', 'P') IS NOT NULL
DROP PROCEDURE [xls].[usp_principal_permissions];
GO
IF OBJECT_ID('[xls].[usp_principal_permissions_change]', 'P') IS NOT NULL
DROP PROCEDURE [xls].[usp_principal_permissions_change];
GO
IF OBJECT_ID('[xls].[usp_role_members]', 'P') IS NOT NULL
DROP PROCEDURE [xls].[usp_role_members];
GO
IF OBJECT_ID('[xls].[usp_role_members_change]', 'P') IS NOT NULL
DROP PROCEDURE [xls].[usp_role_members_change];
GO
IF OBJECT_ID('[xls].[xl_parameter_values_principal]', 'P') IS NOT NULL
DROP PROCEDURE [xls].[xl_parameter_values_principal];
GO
IF OBJECT_ID('[xls].[xl_parameter_values_schema]', 'P') IS NOT NULL
DROP PROCEDURE [xls].[xl_parameter_values_schema];
GO

DECLARE @sql nvarchar(max) = ''

SELECT
    @sql = @sql + 'ALTER ROLE ' + QUOTENAME(r.name) + ' DROP MEMBER ' + QUOTENAME(m.name) + ';' + CHAR(13) + CHAR(10)
FROM
    sys.database_role_members rm
    INNER JOIN sys.database_principals r ON r.principal_id = rm.role_principal_id
    INNER JOIN sys.database_principals m ON m.principal_id = rm.member_principal_id
WHERE
    r.name IN ('xls_admins')

IF LEN(@sql) > 1
    BEGIN
    EXEC (@sql);
    PRINT @sql
    END
GO

IF DATABASE_PRINCIPAL_ID('xls_admins') IS NOT NULL
DROP ROLE [xls_admins];
GO

print 'SaveToDB Administrator Framework removed';
