-- =============================================
-- SaveToDB Administrator Framework for Microsoft SQL Server
-- Version 10.8, January 9, 2023
--
-- This script updates SaveToDB Administrator Framework 9 to the latest version
--
-- Copyright 2018-2023 Gartle LLC
--
-- License: MIT
-- =============================================

IF 1008 <= COALESCE((SELECT CAST(LEFT(HANDLER_CODE, CHARINDEX('.', HANDLER_CODE) - 1) AS int) * 100 + CAST(RIGHT(HANDLER_CODE, LEN(HANDLER_CODE) - CHARINDEX('.', HANDLER_CODE)) AS float) FROM xls.handlers WHERE TABLE_SCHEMA = 'xls' AND TABLE_NAME = 'administrator_framework' AND COLUMN_NAME = 'version' AND EVENT_NAME = 'Information'), 0)
    RAISERROR('SaveToDB Administrator Framework is up-to-date. Update skipped', 11, 0)
GO

-- Add administrators to the xls_users or xls_developers role instead

REVOKE SELECT ON xls.formats        FROM xls_admins;
REVOKE SELECT ON xls.handlers       FROM xls_admins;
REVOKE SELECT ON xls.objects        FROM xls_admins;
REVOKE SELECT ON xls.translations   FROM xls_admins;
REVOKE SELECT ON xls.workbooks      FROM xls_admins;
REVOKE SELECT ON xls.queries        FROM xls_admins;
GO

UPDATE xls.handlers SET EVENT_NAME = 'Information', HANDLER_CODE = '10.8' WHERE TABLE_SCHEMA = 'xls' AND TABLE_NAME = 'administrator_framework' AND COLUMN_NAME = 'version' AND EVENT_NAME = 'Information';
IF @@ROWCOUNT = 0
    INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES ('xls', 'administrator_framework', 'version', 'Information', NULL, NULL, 'ATTRIBUTE', '10.8', NULL, NULL, NULL);
GO

IF (SELECT COUNT(*) FROM xls.handlers WHERE TABLE_SCHEMA = 'xls' AND TABLE_NAME = 'usp_database_permissions' AND EVENT_NAME = 'License') = 0
    BEGIN
    INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_database_permissions', NULL, N'License', NULL, NULL, NULL, N'RugBvtHWd0nZKvfsbNymqeLN283zckC+AftPHaX/8w+xHhQRNuqXqSg7EmazDIj6mTMLeTxx+Izqkdb3961TgWfF5Q8HMIZ1Z+gtPPMO9K4G6SW06Zq/PwKWwlxcjF4Gdz5ZkOTTxUQMC7oEA/3JUGqUY75y9NE4BGEsMbyo+uA=', NULL, NULL, NULL);
    INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_object_permissions', NULL, N'License', NULL, NULL, NULL, N'WHSvJCMuSLwKcFqwcTgbD5U3fu1FeUHDsXKnXpg/ONOuNezMwne3lKPm7aq2rdSkLdA2ZFhkE+azDAJ+XUA/Gia4dPkWldWHMMQh9L4TTz+GuPteC2dfN7BX0c3gyDCAuyLzrxAbsO8+C/y1iZ0Xrz99JSLQ6NCTTbdx8SKmbIU=', NULL, NULL, NULL);
    INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_principal_permissions', NULL, N'License', NULL, NULL, NULL, N'HJujJjKWY+UcBNMELB1XK2UT9SnbEuIUo9DEMdkRJdPCvdl25nVL4ozXMBUywzHVIvsfJLlpToEdWwP3LoS6nS3yweo3bysVAO3egjzHqpAIwf0XhvFErkgJctTo3YlAd9AFS4RuuxiMEUfKhvH0F0WCHV90eOrHSqxuCUiJ6QY=', NULL, NULL, NULL);
    INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_role_members', NULL, N'License', NULL, NULL, NULL, N'1P4SvJTuZQJQpvaQZdzH6MKu23hZqOrhn40V8aQWr+nJlEOCp0LOuzNcmjaH909dzqzBjQy1ruanTqUsVBB+Jm7OMITVuGT57SES7x9HI8HDaYzQbtqeT2xDVvZwIXd4hiM8XVfCDHSNO+F4T4yfGDyuMGjMtdLV/EEQxD/rQOQ=', NULL, NULL, NULL);
    END
GO

UPDATE xls.handlers SET HANDLER_TYPE = N'ATTRIBUTE' WHERE HANDLER_TYPE IS NULL AND EVENT_NAME = 'License';
GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.0, 2022-07-05
-- Description: Selects user roles
-- =============================================

ALTER PROCEDURE [xls].[usp_role_members]
AS
BEGIN

SET NOCOUNT ON;

DECLARE @list varchar(MAX)

SELECT @list = STUFF((
    SELECT ', [' + name + ']' FROM (
        SELECT DISTINCT p.name, p.is_fixed_role FROM sys.database_principals p
            WHERE name NOT IN ('db_owner', 'public') AND p.[type] IN ('R') AND p.is_fixed_role = 0
        ) AS t ORDER BY t.is_fixed_role, t.name
    FOR XML PATH(''), TYPE
    ).value('.', 'nvarchar(MAX)'), 1, 2, '')

IF @list IS NULL
    SELECT
        LOWER(m.type_desc) AS [type]
        , m.name
        , NULL AS format_column
    FROM
        sys.database_principals m
        LEFT JOIN sys.database_role_members rm ON rm.member_principal_id = m.principal_id
        LEFT JOIN sys.database_principals p ON p.principal_id = rm.role_principal_id
    WHERE
        m.[type] IN ('S', 'U', 'R')
        AND m.is_fixed_role = 0
        AND NOT m.name IN ('dbo', 'sys', 'guest', 'public', 'INFORMATION_SCHEMA', 'xls_users', 'xls_developers', 'xls_formats', 'xls_admins', 'doc_readers', 'doc_writers', 'log_app', 'log_admins', 'log_administrators', 'log_users')
    ORDER BY
        m.type_desc
        , m.name
ELSE
    BEGIN
    DECLARE @sql varchar(MAX)
    SET @sql = '
SELECT
    LOWER(p.[type]) AS [type]
    , p.name
    , NULL AS format_column
    , ' + COALESCE(@list, '') + '
    , NULL AS last_format_column
FROM
    (
    SELECT
        p.name AS [role]
        , m.type_desc AS [type]
        , m.name
        , 1 AS [include]
    FROM
        sys.database_principals m
        LEFT JOIN sys.database_role_members rm ON rm.member_principal_id = m.principal_id
        LEFT JOIN sys.database_principals p ON p.principal_id = rm.role_principal_id
    WHERE
        m.[type] IN (''S'', ''U'', ''R'')
        AND m.is_fixed_role = 0
        AND NOT m.name IN (''dbo'', ''sys'', ''guest'', ''public'', ''INFORMATION_SCHEMA'', ''xls_users'', ''xls_developers'', ''xls_formats'', ''xls_admins'', ''doc_readers'', ''doc_writers'', ''log_app'', ''log_admins'', ''log_administrators'', ''log_users'')
    ) s PIVOT (
        SUM([include]) FOR [role] IN ('+ COALESCE(@list, '') + ')
    ) p
ORDER BY
    p.[type]
    , p.[name]
'
    EXEC(@sql)
    -- PRINT @sql
    END

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.0, 2022-07-05
-- Description: Selects database permissions
-- =============================================

ALTER PROCEDURE [xls].[usp_database_permissions]
AS
BEGIN

SET NOCOUNT ON;

WITH cte (
    principal_id
    , [permission_name]
    , [state]
) AS (
    SELECT
        r.principal_id
        , p.[permission_name]
        , p.[state]
    FROM
        sys.database_principals r
        LEFT OUTER JOIN sys.database_permissions p ON p.grantee_principal_id = r.principal_id AND p.class = 0
    WHERE
        r.is_fixed_role = 0
        AND NOT r.[sid] IS NULL
        AND NOT r.name IN ('dbo')
)

SELECT
    LOWER(r.type_desc) AS principal_type
    , r.name AS principal
    , CASE p.[CONNECT]          WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [CONNECT]
    , CASE p.[SELECT]           WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [SELECT]
    , CASE p.[INSERT]           WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [INSERT]
    , CASE p.[UPDATE]           WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [UPDATE]
    , CASE p.[DELETE]           WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [DELETE]
    , CASE p.[EXECUTE]          WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [EXECUTE]
    , CASE p.[VIEW DEFINITION]  WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [VIEW DEFINITION]
    , CASE p.[REFERENCES]       WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [REFERENCES]
    , CASE p.[ALTER ANY SCHEMA] WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [ALTER ANY SCHEMA]
    , CASE p.[CREATE SCHEMA]    WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [CREATE SCHEMA]
    , CASE p.[CREATE TABLE]     WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [CREATE TABLE]
    , CASE p.[CREATE VIEW]      WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [CREATE VIEW]
    , CASE p.[CREATE PROCEDURE] WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [CREATE PROCEDURE]
    , CASE p.[CREATE FUNCTION]  WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [CREATE FUNCTION]
    , CASE p.[ALTER ANY USER]   WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [ALTER ANY USER]
    , CASE p.[ALTER ANY ROLE]   WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [ALTER ANY ROLE]
    , CASE p.[CREATE ROLE]      WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [CREATE ROLE]
    , CASE p.[TAKE OWNERSHIP]   WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [TAKE OWNERSHIP]
    , CASE p.[ALTER]            WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [ALTER]
    , CASE p.[CONTROL]          WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [CONTROL]
FROM
    (
        SELECT
            t.principal_id
            , t.[permission_name]
            , CASE t.[state] WHEN 'D' THEN 43 WHEN 'W' THEN 42 WHEN 'G' THEN 41 ELSE NULL END AS state_mask
        FROM
            cte t
        UNION ALL
        SELECT
            u.principal_id
            , t.[permission_name]
            , CASE t.[state] WHEN 'D' THEN 33 WHEN 'W' THEN 32 WHEN 'G' THEN 31 ELSE NULL END AS state_mask
        FROM
            cte t
            INNER JOIN sys.database_role_members rm ON rm.role_principal_id = t.principal_id
            LEFT OUTER JOIN sys.database_principals u ON u.principal_id = rm.member_principal_id
        WHERE
            t.[permission_name] IS NOT NULL
    ) s PIVOT (
        MAX(state_mask)
        FOR [permission_name] IN (
            [CONNECT]
            , [SELECT], [INSERT], [UPDATE], [DELETE]
            , [EXECUTE]
            , [VIEW DEFINITION]
            , [REFERENCES]
            , [ALTER ANY SCHEMA], [CREATE SCHEMA]
            , [CREATE TABLE], [CREATE VIEW], [CREATE PROCEDURE], [CREATE FUNCTION]
            , [ALTER ANY USER]
            , [ALTER ANY ROLE], [CREATE ROLE]
            , [TAKE OWNERSHIP], [ALTER], [CONTROL]
            )
    ) p
    INNER JOIN sys.database_principals r ON r.principal_id = p.principal_id
ORDER BY
    r.[type]
    , r.name

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.0, 2022-07-05
-- Description: Cell change event handler for usp_database_permissions
-- =============================================

ALTER PROCEDURE [xls].[usp_database_permissions_change]
    @column_name nvarchar(128) = NULL
    , @cell_value int = NULL
    , @principal nvarchar(128) = NULL
AS
BEGIN

SET NOCOUNT ON

DECLARE @permission varchar(128)
DECLARE @authorization varchar(128)
DECLARE @options varchar(128) = ''

SET @permission = UPPER(@column_name)
SET @authorization = UPPER(@cell_value)

IF CHARINDEX('-' + @permission + '-',
    '-CONNECT-SELECT-INSERT-UPDATE-DELETE-EXECUTE-VIEW DEFINITION-REFERENCES-ALTER ANY SCHEMA-CREATE SCHEMA-CREATE TABLE-CREATE VIEW-CREATE PROCEDURE-CREATE FUNCTION-ALTER ANY USER-ALTER ANY ROLE-CREATE ROLE-TAKE OWNERSHIP-ALTER-CONTROL-') = 0
    RETURN

IF @authorization IS NULL
    SET @authorization = 'REVOKE'
ELSE IF @authorization = 'GRANT' OR @authorization = 'G'
    SET @authorization = 'GRANT'
ELSE IF @authorization = 'GRANT+' OR @authorization = 'G+'
    BEGIN
    SET @authorization = 'GRANT'
    SET @options      = ' WITH GRANT OPTION'
    END
ELSE IF @authorization = 'DENY' OR @authorization = 'D'
    SET @authorization = 'DENY'
ELSE IF @authorization = 'DENY+' OR @authorization = 'D+'
    BEGIN
    SET @authorization = 'DENY'
    SET @options      = ' CASCADE'
    END
ELSE IF @authorization = 'REVOKE' OR @authorization = 'R'
    SET @authorization = 'REVOKE'
ELSE IF @authorization = 'REVOKE-' OR @authorization = 'R-'
    SET @authorization = 'REVOKE GRANT OPTION FOR'
ELSE IF @authorization = 'REVOKE+' OR @authorization = 'R+'
    BEGIN
    SET @authorization = 'REVOKE'
    SET @options      = ' CASCADE'
    END
ELSE
    RETURN

DECLARE @query varchar(max)

SET @query = @authorization + ' ' + @permission + ' TO ' + QUOTENAME(@principal) + @options

EXEC (@query)

END


GO

print 'SaveToDB Administrator Framework updated';
