-- =============================================
-- SaveToDB Administrator Framework for Microsoft SQL Server
-- Version 10.8, January 9, 2023
--
-- Copyright 2018-2023 Gartle LLC
--
-- License: MIT
-- =============================================

SET NOCOUNT ON
GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.0, 2022-07-05
-- Description: Selects database permissions
-- =============================================

CREATE PROCEDURE [xls].[usp_database_permissions]
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

CREATE PROCEDURE [xls].[usp_database_permissions_change]
    @column_name nvarchar(128) = NULL
    , @cell_value nvarchar(128) = NULL
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

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.0, 2022-07-05
-- Description: Selects object permissions
-- =============================================

CREATE PROCEDURE [xls].[usp_object_permissions]
    @principal nvarchar(128) = NULL,
    @schema nvarchar(128) = NULL,
    @type nvarchar(128) = NULL,
    @has_any bit = NULL,
    @has_direct bit = NULL
AS
BEGIN

IF @type = '' SET @type = NULL

SET NOCOUNT ON;

WITH cte (
    [schema_id]
    , [object_id]
    , principal_id
    , [permission_name]
    , [state]
) AS (
    SELECT
        o.[schema_id]
        , o.[object_id]
        , r.principal_id
        , p.[permission_name]
        , p.[state]
    FROM
        sys.objects o
        INNER JOIN sys.schemas s ON s.[schema_id] = o.[schema_id]
        LEFT OUTER JOIN sys.extended_properties e ON e.major_id = o.[object_id]
        CROSS JOIN sys.database_principals r
        LEFT OUTER JOIN sys.database_permissions p ON o.[object_id] = p.major_id AND r.principal_id = p.grantee_principal_id AND p.class = 1
    WHERE
        o.[type] IN ('U', 'V', 'P', 'FN', 'IF')
        AND ISNULL(o.is_ms_shipped, 0) = 0
        AND NOT ISNULL(e.name, '') = 'microsoft_database_tools_support'
        AND r.is_fixed_role = 0
        AND NOT r.[sid] IS NULL
        AND NOT r.name IN ('dbo', 'guest', 'public')
    UNION ALL
    SELECT
        s.[schema_id]
        , NULL AS [object_id]
        , r.principal_id
        , p.[permission_name]
        , p.[state]
    FROM
        sys.schemas s
        INNER JOIN sys.database_principals e ON e.principal_id = s.principal_id
        CROSS JOIN sys.database_principals r
        LEFT OUTER JOIN sys.database_permissions p ON p.major_id = s.[schema_id] AND r.principal_id = p.grantee_principal_id AND p.class = 3
    WHERE
        r.is_fixed_role = 0
        AND NOT r.[sid] IS NULL
        AND NOT r.name IN ('dbo', 'guest', 'public')
        AND e.is_fixed_role = 0
        AND NOT e.[sid] IS NULL
        AND NOT s.name IN ('sys', 'guest', 'INFORMATION_SCHEMA')
    )

SELECT
    t.principal_type
    , t.principal
    , t.[schema]
    , t.name
    , t.[type]
    , t.has_any
    , t.has_direct
    , t.[SELECT]
    , t.[INSERT]
    , t.[UPDATE]
    , t.[DELETE]
    , t.[EXECUTE]
    , t.[VIEW DEFINITION]
    , t.[REFERENCES], [TAKE OWNERSHIP]
    , t.[ALTER]
    , t.[CONTROL]
FROM
    (
        SELECT
            LOWER(r.type_desc) AS principal_type
            , r.name AS principal
            , s.name AS [schema]
            , o.name
            , CASE o.[type]
                WHEN 'U'  THEN 'table'
                WHEN 'V'  THEN 'view'
                WHEN 'P'  THEN 'procedure'
                WHEN 'IF' THEN 'function'
                WHEN 'FN' THEN 'function'
                ELSE COALESCE(LOWER(o.type_desc), 'schema')
                END AS [type]
            , CASE WHEN COALESCE([SELECT], [INSERT], [UPDATE], [DELETE], [EXECUTE], [VIEW DEFINITION], [REFERENCES], [TAKE OWNERSHIP], [CONTROL], 0) = 0 THEN 0 ELSE 1 END AS has_any
            , CASE p.[SELECT]           WHEN 43 THEN 1 WHEN 42 THEN 1 WHEN 41 THEN 1 ELSE
              CASE p.[INSERT]           WHEN 43 THEN 1 WHEN 42 THEN 1 WHEN 41 THEN 1 ELSE
              CASE p.[UPDATE]           WHEN 43 THEN 1 WHEN 42 THEN 1 WHEN 41 THEN 1 ELSE
              CASE p.[DELETE]           WHEN 43 THEN 1 WHEN 42 THEN 1 WHEN 41 THEN 1 ELSE
              CASE p.[EXECUTE]          WHEN 43 THEN 1 WHEN 42 THEN 1 WHEN 41 THEN 1 ELSE
              CASE p.[VIEW DEFINITION]  WHEN 43 THEN 1 WHEN 42 THEN 1 WHEN 41 THEN 1 ELSE
              CASE p.[REFERENCES]       WHEN 43 THEN 1 WHEN 42 THEN 1 WHEN 41 THEN 1 ELSE
              CASE p.[TAKE OWNERSHIP]   WHEN 43 THEN 1 WHEN 42 THEN 1 WHEN 41 THEN 1 ELSE
              CASE p.[ALTER]            WHEN 43 THEN 1 WHEN 42 THEN 1 WHEN 41 THEN 1 ELSE
              CASE p.[CONTROL]          WHEN 43 THEN 1 WHEN 42 THEN 1 WHEN 41 THEN 1 ELSE
                0 END END END END END END END END END END AS has_direct
            , CASE p.[SELECT]           WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' WHEN 23 THEN 'DENY s' WHEN 22 THEN 'GRANT+ s' WHEN 21 THEN 'GRANT s' WHEN 13 THEN 'DENY sr' WHEN 12 THEN 'GRANT+ sr' WHEN 11 THEN 'GRANT sr' ELSE NULL END AS [SELECT]
            , CASE p.[INSERT]           WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' WHEN 23 THEN 'DENY s' WHEN 22 THEN 'GRANT+ s' WHEN 21 THEN 'GRANT s' WHEN 13 THEN 'DENY sr' WHEN 12 THEN 'GRANT+ sr' WHEN 11 THEN 'GRANT sr' ELSE NULL END AS [INSERT]
            , CASE p.[UPDATE]           WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' WHEN 23 THEN 'DENY s' WHEN 22 THEN 'GRANT+ s' WHEN 21 THEN 'GRANT s' WHEN 13 THEN 'DENY sr' WHEN 12 THEN 'GRANT+ sr' WHEN 11 THEN 'GRANT sr' ELSE NULL END AS [UPDATE]
            , CASE p.[DELETE]           WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' WHEN 23 THEN 'DENY s' WHEN 22 THEN 'GRANT+ s' WHEN 21 THEN 'GRANT s' WHEN 13 THEN 'DENY sr' WHEN 12 THEN 'GRANT+ sr' WHEN 11 THEN 'GRANT sr' ELSE NULL END AS [DELETE]
            , CASE p.[EXECUTE]          WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' WHEN 23 THEN 'DENY s' WHEN 22 THEN 'GRANT+ s' WHEN 21 THEN 'GRANT s' WHEN 13 THEN 'DENY sr' WHEN 12 THEN 'GRANT+ sr' WHEN 11 THEN 'GRANT sr' ELSE NULL END AS [EXECUTE]
            , CASE p.[VIEW DEFINITION]  WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' WHEN 23 THEN 'DENY s' WHEN 22 THEN 'GRANT+ s' WHEN 21 THEN 'GRANT s' WHEN 13 THEN 'DENY sr' WHEN 12 THEN 'GRANT+ sr' WHEN 11 THEN 'GRANT sr' ELSE NULL END AS [VIEW DEFINITION]
            , CASE p.[REFERENCES]       WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' WHEN 23 THEN 'DENY s' WHEN 22 THEN 'GRANT+ s' WHEN 21 THEN 'GRANT s' WHEN 13 THEN 'DENY sr' WHEN 12 THEN 'GRANT+ sr' WHEN 11 THEN 'GRANT sr' ELSE NULL END AS [REFERENCES]
            , CASE p.[TAKE OWNERSHIP]   WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' WHEN 23 THEN 'DENY s' WHEN 22 THEN 'GRANT+ s' WHEN 21 THEN 'GRANT s' WHEN 13 THEN 'DENY sr' WHEN 12 THEN 'GRANT+ sr' WHEN 11 THEN 'GRANT sr' ELSE NULL END AS [TAKE OWNERSHIP]
            , CASE p.[ALTER]            WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' WHEN 23 THEN 'DENY s' WHEN 22 THEN 'GRANT+ s' WHEN 21 THEN 'GRANT s' WHEN 13 THEN 'DENY sr' WHEN 12 THEN 'GRANT+ sr' WHEN 11 THEN 'GRANT sr' ELSE NULL END AS [ALTER]
            , CASE p.[CONTROL]          WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' WHEN 23 THEN 'DENY s' WHEN 22 THEN 'GRANT+ s' WHEN 21 THEN 'GRANT s' WHEN 13 THEN 'DENY sr' WHEN 12 THEN 'GRANT+ sr' WHEN 11 THEN 'GRANT sr' ELSE NULL END AS [CONTROL]
            , CASE o.[type]
                WHEN 'U'  THEN 11
                WHEN 'V'  THEN 12
                WHEN 'P'  THEN 13
                WHEN 'IF' THEN 14
                WHEN 'FN' THEN 15
                ELSE CASE WHEN o.[type] IS NULL THEN 1 ELSE 21 END END AS type_order
        FROM
            (
                SELECT
                    t.[schema_id]
                    , t.[object_id]
                    , t.principal_id
                    , t.[permission_name]
                    , CASE t.[state] WHEN 'D' THEN 43 WHEN 'W' THEN 42 WHEN 'G' THEN 41 ELSE NULL END AS state_mask
                FROM
                    cte t
                UNION ALL
                SELECT
                    t.[schema_id]
                    , t.[object_id]
                    , u.principal_id
                    , t.[permission_name]
                    , CASE t.[state] WHEN 'D' THEN 33 WHEN 'W' THEN 32 WHEN 'G' THEN 31 ELSE NULL END AS state_mask
                FROM
                    cte t
                    INNER JOIN sys.database_role_members rm ON rm.role_principal_id = t.principal_id
                    LEFT OUTER JOIN sys.database_principals u ON u.principal_id = rm.member_principal_id
                WHERE
                    t.[permission_name] IS NOT NULL
                UNION ALL
                SELECT
                    t.[schema_id]
                    , o.[object_id]
                    , t.principal_id
                    , t.[permission_name]
                    , CASE t.[state] WHEN 'D' THEN 23 WHEN 'W' THEN 22 WHEN 'G' THEN 21 ELSE NULL END AS state_mask
                FROM
                    cte t
                    INNER JOIN sys.objects o ON o.[schema_id] = t.[schema_id]
                WHERE
                    t.[permission_name] IS NOT NULL
                    AND t.[object_id] IS NULL
                    AND o.[type] IN ('U', 'V', 'P', 'FN', 'IF')
                UNION ALL
                SELECT
                    t.[schema_id]
                    , o.[object_id]
                    , u.principal_id
                    , t.[permission_name]
                    , CASE t.[state] WHEN 'D' THEN 13 WHEN 'W' THEN 12 WHEN 'G' THEN 11 ELSE NULL END AS state_mask
                FROM
                    cte t
                    INNER JOIN sys.objects o ON o.[schema_id] = t.[schema_id]
                    INNER JOIN sys.database_role_members rm ON rm.role_principal_id = t.principal_id
                    LEFT OUTER JOIN sys.database_principals u ON u.principal_id = rm.member_principal_id
                WHERE
                    t.[permission_name] IS NOT NULL
                    AND t.[object_id] IS NULL
                    AND o.[type] IN ('U', 'V', 'P', 'FN', 'IF')
            ) s PIVOT (
                MAX(state_mask)
                FOR [permission_name] IN (
                    [SELECT], [INSERT], [UPDATE], [DELETE]
                    , [EXECUTE]
                    , [VIEW DEFINITION]
                    , [REFERENCES], [TAKE OWNERSHIP], [ALTER]
                    , [CONTROL]
                    )
            ) p
            INNER JOIN sys.schemas s ON s.[schema_id] = p.[schema_id]
            LEFT OUTER JOIN sys.objects o ON o.[object_id] = p.[object_id]
            INNER JOIN sys.database_principals r ON r.principal_id = p.principal_id
    ) t
WHERE
    COALESCE(t.principal, '') = COALESCE(@principal, t.principal, '')
    AND COALESCE(t.[schema], '') = COALESCE(@schema, t.[schema], '')
    AND COALESCE(t.[type], '') = COALESCE(@type, t.[type], '')
    AND COALESCE(t.has_any, -1) = COALESCE(CAST(@has_any AS int), t.has_any, -1)
    AND COALESCE(t.has_direct, -1) = COALESCE(CAST(@has_direct AS int), t.has_direct, -1)
ORDER BY
    t.principal_type
    , t.principal
    , t.[schema]
    , t.type_order
    , t.[type]
    , t.name

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.0, 2022-07-05
-- Description: Cell change event handler for usp_object_permissions
-- =============================================

CREATE PROCEDURE [xls].[usp_object_permissions_change]
    @column_name nvarchar(128) = NULL
    , @cell_value nvarchar(128) = NULL
    , @principal nvarchar(128) = NULL
    , @schema nvarchar(128) = NULL
    , @name nvarchar(128) = NULL
AS
BEGIN

SET NOCOUNT ON

DECLARE @permission varchar(128)
DECLARE @authorization varchar(128)

SET @permission = UPPER(@column_name)
SET @authorization = UPPER(@cell_value)
DECLARE @options varchar(128) = ''

IF CHARINDEX('-' + @permission + '-',
    '-SELECT-INSERT-UPDATE-DELETE-EXECUTE-VIEW DEFINITION-REFERENCES-TAKE OWNERSHIP-ALTER-CONTROL-') = 0
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

IF @name IS NULL
    SET @query = @authorization + ' ' + @permission + ' ON SCHEMA::' + QUOTENAME(@schema) + ' TO ' + QUOTENAME(@principal)
ELSE
    SET @query = @authorization + ' ' + @permission + ' ON OBJECT::' +
        QUOTENAME(@schema) + '.' + QUOTENAME(@name) + ' TO ' + QUOTENAME(@principal)

EXEC (@query)

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.0, 2022-07-05
-- Description: Selects object permissions
-- =============================================

CREATE PROCEDURE [xls].[usp_principal_permissions]
    @principal nvarchar(128) = NULL,
    @name nvarchar(128) = NULL,
    @has_any bit = NULL
AS
BEGIN

SET NOCOUNT ON;

WITH cte (
    principal_id
    , [user_id]
    , [permission_name]
    , [state]
) AS (
    SELECT
        r.principal_id
        , u.principal_id AS [user_id]
        , p.[permission_name]
        , p.[state]
    FROM
        sys.database_principals u
        CROSS JOIN sys.database_principals r
        LEFT OUTER JOIN sys.database_permissions p ON p.major_id = u.principal_id AND r.principal_id = p.grantee_principal_id AND p.class = 4
    WHERE
        r.is_fixed_role = 0
        AND NOT r.[sid] IS NULL
        AND NOT r.name IN ('guest', 'public')
        AND u.is_fixed_role = 0
        AND NOT u.[sid] IS NULL
        AND NOT u.name IN ('dbo', 'guest', 'public')
        AND NOT r.principal_id = u.principal_id
)

SELECT
    t.principal_type
    , t.principal
    , t.name
    , t.[type]
    , t.has_any
    , t.[VIEW DEFINITION]
    , t.[IMPERSONATE]
    , t.[TAKE OWNERSHIP]
    , t.[ALTER]
    , t.[CONTROL]
FROM
    (
        SELECT
            LOWER(r.type_desc) AS principal_type
            , r.name AS principal
            , u.name AS name
            , LOWER(u.type_desc) AS [type]
            , CASE WHEN COALESCE([VIEW DEFINITION], [IMPERSONATE], [TAKE OWNERSHIP], [ALTER], [CONTROL], 0) = 0 THEN 0 ELSE 1 END AS has_any
            , CASE p.[VIEW DEFINITION]  WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [VIEW DEFINITION]
            , CASE p.[IMPERSONATE]      WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [IMPERSONATE]
            , CASE p.[TAKE OWNERSHIP]   WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [TAKE OWNERSHIP]
            , CASE p.[ALTER]            WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [ALTER]
            , CASE p.[CONTROL]          WHEN 43 THEN 'DENY' WHEN 42 THEN 'GRANT+' WHEN 41 THEN 'GRANT' WHEN 33 THEN 'DENY r' WHEN 32 THEN 'GRANT+ r' WHEN 31 THEN 'GRANT r' ELSE NULL END AS [CONTROL]
        FROM
            (
                SELECT
                    t.principal_id
                    , t.[user_id]
                    , t.[permission_name]
                    , CASE t.[state] WHEN 'D' THEN 43 WHEN 'W' THEN 42 WHEN 'G' THEN 41 ELSE NULL END AS state_mask
                FROM
                    cte t
                UNION ALL
                SELECT
                    u.principal_id
                    , t.[user_id]
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
                    [VIEW DEFINITION]
                    , [IMPERSONATE]
                    , [TAKE OWNERSHIP], [ALTER], [CONTROL]
                    )
            ) p
            INNER JOIN sys.database_principals r ON r.principal_id = p.principal_id
            INNER JOIN sys.database_principals u ON u.principal_id = p.[user_id]
    ) t
WHERE
    COALESCE(t.principal, '') = COALESCE(@principal, t.principal, '')
    AND COALESCE(t.name, '') = COALESCE(@name, t.name, '')
    AND COALESCE(t.has_any, -1) = COALESCE(@has_any, t.has_any, -1)
ORDER BY
    t.principal_type
    , t.principal
    , t.[type]
    , t.name

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.0, 2022-07-05
-- Description: Cell change event handler for usp_object_permissions
-- =============================================

CREATE PROCEDURE [xls].[usp_principal_permissions_change]
    @column_name nvarchar(128) = NULL
    , @cell_value int = NULL
    , @principal nvarchar(128) = NULL
    , @name nvarchar(128) = NULL
    , @type nvarchar(128) = NULL
AS
BEGIN

SET NOCOUNT ON

DECLARE @permission varchar(128)
DECLARE @authorization varchar(128)

SET @permission = UPPER(@column_name)
SET @authorization = UPPER(@cell_value)
DECLARE @options varchar(128) = ''

-- Start from here

SET NOCOUNT ON

IF CHARINDEX('-' + @permission + '-',
    '-VIEW DEFINITION-IMPERSONATE-TAKE OWNERSHIP-ALTER-CONTROL-') = 0
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

IF CHARINDEX('user', @type) > 0
    SET @query = @authorization + ' ' + @permission + ' ON USER::' + QUOTENAME(@name) + ' TO ' + QUOTENAME(@principal) + @options
ELSE IF CHARINDEX('application', @type) > 0
    SET @query = @authorization + ' ' + @permission + ' ON APPLICATION ROLE::' + QUOTENAME(@name) + ' TO ' + QUOTENAME(@principal) + @options
ELSE
    SET @query = @authorization + ' ' + @permission + ' ON ROLE::' + QUOTENAME(@name) + ' TO ' + QUOTENAME(@principal) + @options

EXEC (@query)

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.0, 2022-07-05
-- Description: Selects user roles
-- =============================================

CREATE PROCEDURE [xls].[usp_role_members]
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
-- Description: Cell change event handler for usp_roles
-- =============================================

CREATE PROCEDURE [xls].[usp_role_members_change]
    @column_name nvarchar(128) = NULL
    , @cell_number_value int = NULL
    , @name nvarchar(128) = NULL
    , @data_language varchar(10) = NULL
AS
BEGIN

SET NOCOUNT ON

DECLARE @message nvarchar(max)

IF DATABASE_PRINCIPAL_ID(@column_name) IS NULL
    BEGIN
    SET @message = N'Role ''%s'' does not exist'
    SET @message = COALESCE((SELECT TOP 1 TRANSLATED_NAME FROM xls.translations WHERE TABLE_SCHEMA = 'xls' AND TABLE_NAME = 'strings' AND LANGUAGE_NAME =  @data_language AND COLUMN_NAME = @message), @message)
    RAISERROR(@message, 11, 0, @column_name);
    RETURN
    END

IF DATABASE_PRINCIPAL_ID(@name) IS NULL
    BEGIN
    SET @message = N'User ''%s'' does not exist'
    SET @message = COALESCE((SELECT TOP 1 TRANSLATED_NAME FROM xls.translations WHERE TABLE_SCHEMA = 'xls' AND TABLE_NAME = 'strings' AND LANGUAGE_NAME =  @data_language AND COLUMN_NAME = @message), @message)
    RAISERROR(@message, 11, 0, @name);
    RETURN
    END

IF @cell_number_value = 0
    EXEC sp_droprolemember @column_name, @name
ELSE IF @cell_number_value = 1
    EXEC sp_addrolemember @column_name, @name
ELSE
    BEGIN
    SET @message = N'Set 1 to add and 0 to remove a user from the role'
    SET @message = COALESCE((SELECT TOP 1 TRANSLATED_NAME FROM xls.translations WHERE TABLE_SCHEMA = 'xls' AND TABLE_NAME = 'strings' AND LANGUAGE_NAME =  @data_language AND COLUMN_NAME = @message), @message)
    RAISERROR(@message, 11, 0);
    RETURN
    END

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.0, 2022-07-05
-- Description: The procedure selects values for the @principal parameter
-- =============================================

CREATE PROCEDURE [xls].[xl_parameter_values_principal]
AS
BEGIN

SET NOCOUNT ON

SELECT NULL AS name UNION ALL
SELECT
    name
FROM
    sys.database_principals r
WHERE
    r.is_fixed_role = 0
    AND NOT r.name IN ('dbo', 'sys', 'public', 'guest', 'INFORMATION_SCHEMA')
ORDER BY
    name

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.0, 2022-07-05
-- Description: The procedure selects values for the @schame parameter
-- =============================================

CREATE PROCEDURE [xls].[xl_parameter_values_schema]
AS
BEGIN

SET NOCOUNT ON

SELECT NULL AS name UNION ALL
SELECT
    s.name
FROM
    sys.schemas s
    INNER JOIN sys.database_principals e ON e.principal_id = s.principal_id
WHERE
    e.is_fixed_role = 0
    AND NOT e.[sid] IS NULL
    AND NOT s.name IN ('sys', 'guest', 'INFORMATION_SCHEMA')
ORDER BY
    name

END


GO

CREATE ROLE xls_admins;
GO

GRANT VIEW DEFINITION ON SCHEMA::xls TO xls_admins;

GRANT VIEW DEFINITION ON ROLE::xls_users      TO xls_admins;
GRANT VIEW DEFINITION ON ROLE::xls_formats    TO xls_admins;
GRANT VIEW DEFINITION ON ROLE::xls_developers TO xls_admins;

GRANT EXECUTE ON xls.usp_database_permissions         TO xls_admins;
GRANT EXECUTE ON xls.usp_database_permissions_change  TO xls_admins;
GRANT EXECUTE ON xls.usp_object_permissions           TO xls_admins;
GRANT EXECUTE ON xls.usp_object_permissions_change    TO xls_admins;
GRANT EXECUTE ON xls.usp_principal_permissions        TO xls_admins;
GRANT EXECUTE ON xls.usp_principal_permissions_change TO xls_admins;
GRANT EXECUTE ON xls.usp_role_members                 TO xls_admins;
GRANT EXECUTE ON xls.usp_role_members_change          TO xls_admins;
GRANT EXECUTE ON xls.xl_parameter_values_principal    TO xls_admins;
GRANT EXECUTE ON xls.xl_parameter_values_schema       TO xls_admins;
GO

INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'administrator_framework', N'version', N'Information', NULL, NULL, N'ATTRIBUTE', N'10.0', NULL, NULL, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_object_permissions', N'principal', N'ParameterValues', N'xls', N'xl_parameter_values_principal', N'PROCEDURE', NULL, NULL, NULL, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_object_permissions', N'schema', N'ParameterValues', N'xls', N'xl_parameter_values_schema', N'PROCEDURE', NULL, NULL, NULL, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_object_permissions', N'type', N'ParameterValues', N'xls', N'xl_parameter_values_type', N'VALUES', N', schema, table, view, procedure, function', NULL, NULL, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_principal_permissions', N'principal', N'ParameterValues', N'xls', N'xl_parameter_values_principal', N'PROCEDURE', NULL, NULL, NULL, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_principal_permissions', N'name', N'ParameterValues', N'xls', N'xl_parameter_values_principal', N'PROCEDURE', NULL, NULL, NULL, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_database_permissions', NULL, N'Actions', N'xls', N'SaveToDB Framework Online Help', N'HTTP', N'https://www.savetodb.com/help/savetodb-framework-procedures.htm#xls.usp_database_permissions', NULL, 11, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_database_permissions', NULL, N'Actions', N'xls', N'MenuSeparator20', N'MENUSEPARATOR', NULL, NULL, 20, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_database_permissions', NULL, N'Actions', N'xls', N'Transact-SQL Reference - Permissions', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms191291.aspx', NULL, 21, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_database_permissions', NULL, N'Actions', N'xls', N'Transact-SQL Reference - GRANT Database Permissions', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms187940.aspx', NULL, 22, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_database_permissions', NULL, N'Actions', N'xls', N'Transact-SQL Reference - DENY Database Permissions', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms177518.aspx', NULL, 23, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_database_permissions', NULL, N'Actions', N'xls', N'Transact-SQL Reference - REVOKE Database Permissions', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms177573.aspx', NULL, 24, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_object_permissions', NULL, N'Actions', N'xls', N'SaveToDB Framework Online Help', N'HTTP', N'https://www.savetodb.com/help/savetodb-framework-procedures.htm#xls.usp_object_permissions', NULL, 11, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_object_permissions', NULL, N'Actions', N'xls', N'MenuSeparator20', N'MENUSEPARATOR', NULL, NULL, 20, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_object_permissions', NULL, N'Actions', N'xls', N'Transact-SQL Reference - Permissions', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms191291.aspx', NULL, 21, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_object_permissions', NULL, N'Actions', N'xls', N'Transact-SQL Reference - GRANT Schema Permissions', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms187940.aspx', NULL, 22, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_object_permissions', NULL, N'Actions', N'xls', N'Transact-SQL Reference - GRANT Object Permissions', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms188371.aspx', NULL, 23, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_object_permissions', NULL, N'Actions', N'xls', N'Transact-SQL Reference - DENY Schema Permissions', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms176128.aspx', NULL, 24, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_object_permissions', NULL, N'Actions', N'xls', N'Transact-SQL Reference - DENY Object Permissions', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms173724.aspx', NULL, 25, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_object_permissions', NULL, N'Actions', N'xls', N'Transact-SQL Reference - REVOKE Schema Permissions', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms187733.aspx', NULL, 26, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_object_permissions', NULL, N'Actions', N'xls', N'Transact-SQL Reference - REVOKE Object Permissions', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms187719.aspx', NULL, 27, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_principal_permissions', NULL, N'Actions', N'xls', N'SaveToDB Framework Online Help', N'HTTP', N'https://www.savetodb.com/help/savetodb-framework-procedures.htm#xls.usp_principal_permissions', NULL, 11, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_principal_permissions', NULL, N'Actions', N'xls', N'MenuSeparator20', N'MENUSEPARATOR', NULL, NULL, 20, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_principal_permissions', NULL, N'Actions', N'xls', N'Transact-SQL Reference - Permissions', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms191291.aspx', NULL, 21, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_principal_permissions', NULL, N'Actions', N'xls', N'Transact-SQL Reference - GRANT Database Principal Permissions', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms173848.aspx', NULL, 22, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_principal_permissions', NULL, N'Actions', N'xls', N'Transact-SQL Reference - DENY Database Principal Permissions', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms177518.aspx', NULL, 23, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_principal_permissions', NULL, N'Actions', N'xls', N'Transact-SQL Reference - REVOKE Database Principal Permissions', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms188356.aspx', NULL, 24, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_role_members', NULL, N'Actions', N'xls', N'SaveToDB Framework Online Help', N'HTTP', N'https://www.savetodb.com/help/savetodb-framework-procedures.htm#xls.usp_role_members', NULL, 11, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_role_members', NULL, N'Actions', N'xls', N'MenuSeparator20', N'MENUSEPARATOR', NULL, NULL, 20, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_role_members', NULL, N'Actions', N'xls', N'Transact-SQL Reference - Permissions', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms191291.aspx', NULL, 21, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_role_members', NULL, N'Actions', N'xls', N'Transact-SQL Reference - CREATE ROLE', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms187936.aspx', NULL, 22, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_role_members', NULL, N'Actions', N'xls', N'Transact-SQL Reference - ALTER ROLE', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms189775.aspx', NULL, 23, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_role_members', NULL, N'Actions', N'xls', N'Transact-SQL Reference - DROP ROLE', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms174988.aspx', NULL, 24, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_role_members', NULL, N'Actions', N'xls', N'Transact-SQL Reference - sp_addrolemember', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms187750.aspx', NULL, 25, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_role_members', NULL, N'Actions', N'xls', N'Transact-SQL Reference - sp_droprolemember', N'HTTP', N'https://msdn.microsoft.com/en-us/library/ms188369.aspx', NULL, 26, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_database_permissions', NULL, N'License', NULL, NULL, N'ATTRIBUTE', N'RugBvtHWd0nZKvfsbNymqeLN283zckC+AftPHaX/8w+xHhQRNuqXqSg7EmazDIj6mTMLeTxx+Izqkdb3961TgWfF5Q8HMIZ1Z+gtPPMO9K4G6SW06Zq/PwKWwlxcjF4Gdz5ZkOTTxUQMC7oEA/3JUGqUY75y9NE4BGEsMbyo+uA=', NULL, NULL, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_object_permissions', NULL, N'License', NULL, NULL, N'ATTRIBUTE', N'WHSvJCMuSLwKcFqwcTgbD5U3fu1FeUHDsXKnXpg/ONOuNezMwne3lKPm7aq2rdSkLdA2ZFhkE+azDAJ+XUA/Gia4dPkWldWHMMQh9L4TTz+GuPteC2dfN7BX0c3gyDCAuyLzrxAbsO8+C/y1iZ0Xrz99JSLQ6NCTTbdx8SKmbIU=', NULL, NULL, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_principal_permissions', NULL, N'License', NULL, NULL, N'ATTRIBUTE', N'HJujJjKWY+UcBNMELB1XK2UT9SnbEuIUo9DEMdkRJdPCvdl25nVL4ozXMBUywzHVIvsfJLlpToEdWwP3LoS6nS3yweo3bysVAO3egjzHqpAIwf0XhvFErkgJctTo3YlAd9AFS4RuuxiMEUfKhvH0F0WCHV90eOrHSqxuCUiJ6QY=', NULL, NULL, NULL);
INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_role_members', NULL, N'License', NULL, NULL, N'ATTRIBUTE', N'1P4SvJTuZQJQpvaQZdzH6MKu23hZqOrhn40V8aQWr+nJlEOCp0LOuzNcmjaH909dzqzBjQy1ruanTqUsVBB+Jm7OMITVuGT57SES7x9HI8HDaYzQbtqeT2xDVvZwIXd4hiM8XVfCDHSNO+F4T4yfGDyuMGjMtdLV/EEQxD/rQOQ=', NULL, NULL, NULL);
GO

INSERT INTO xls.formats (TABLE_SCHEMA, TABLE_NAME, TABLE_EXCEL_FORMAT_XML) VALUES (N'xls', N'usp_database_permissions', N'<table name="xls.usp_database_permissions"><columnFormats><column name="" property="ListObjectName" value="database_permissions" type="String"/><column name="" property="ShowTotals" value="False" type="Boolean"/><column name="" property="TableStyle.Name" value="TableStyleMedium15" type="String"/><column name="" property="ShowTableStyleColumnStripes" value="False" type="Boolean"/><column name="" property="ShowTableStyleFirstColumn" value="False" type="Boolean"/><column name="" property="ShowShowTableStyleLastColumn" value="False" type="Boolean"/><column name="" property="ShowTableStyleRowStripes" value="False" type="Boolean"/><column name="_RowNum" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="_RowNum" property="Address" value="$B$4" type="String"/><column name="_RowNum" property="NumberFormat" value="General" type="String"/><column name="principal_type" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="principal_type" property="Address" value="$C$4" type="String"/><column name="principal_type" property="ColumnWidth" value="15.71" type="Double"/><column name="principal_type" property="NumberFormat" value="General" type="String"/><column name="principal" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="principal" property="Address" value="$D$4" type="String"/><column name="principal" property="ColumnWidth" value="26.71" type="Double"/><column name="principal" property="NumberFormat" value="General" type="String"/><column name="CONNECT" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="CONNECT" property="Address" value="$E$4" type="String"/><column name="CONNECT" property="ColumnWidth" value="11.14" type="Double"/><column name="CONNECT" property="NumberFormat" value="General" type="String"/><column name="SELECT" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="SELECT" property="Address" value="$F$4" type="String"/><column name="SELECT" property="ColumnWidth" value="8.57" type="Double"/><column name="SELECT" property="NumberFormat" value="General" type="String"/><column name="INSERT" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="INSERT" property="Address" value="$G$4" type="String"/><column name="INSERT" property="ColumnWidth" value="8.71" type="Double"/><column name="INSERT" property="NumberFormat" value="General" type="String"/><column name="UPDATE" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="UPDATE" property="Address" value="$H$4" type="String"/><column name="UPDATE" property="ColumnWidth" value="9.71" type="Double"/><column name="UPDATE" property="NumberFormat" value="General" type="String"/><column name="DELETE" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="DELETE" property="Address" value="$I$4" type="String"/><column name="DELETE" property="ColumnWidth" value="8.71" type="Double"/><column name="DELETE" property="NumberFormat" value="General" type="String"/><column name="EXECUTE" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="EXECUTE" property="Address" value="$J$4" type="String"/><column name="EXECUTE" property="ColumnWidth" value="10.29" type="Double"/><column name="EXECUTE" property="NumberFormat" value="General" type="String"/><column name="VIEW DEFINITION" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="VIEW DEFINITION" property="Address" value="$K$4" type="String"/><column name="VIEW DEFINITION" property="ColumnWidth" value="18.29" type="Double"/><column name="VIEW DEFINITION" property="NumberFormat" value="General" type="String"/><column name="REFERENCES" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="REFERENCES" property="Address" value="$L$4" type="String"/><column name="REFERENCES" property="ColumnWidth" value="13.43" type="Double"/><column name="REFERENCES" property="NumberFormat" value="General" type="String"/><column name="ALTER ANY SCHEMA" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="ALTER ANY SCHEMA" property="Address" value="$M$4" type="String"/><column name="ALTER ANY SCHEMA" property="ColumnWidth" value="20.43" type="Double"/><column name="ALTER ANY SCHEMA" property="NumberFormat" value="General" type="String"/><column name="CREATE SCHEMA" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="CREATE SCHEMA" property="Address" value="$N$4" type="String"/><column name="CREATE SCHEMA" property="ColumnWidth" value="17.29" type="Double"/><column name="CREATE SCHEMA" property="NumberFormat" value="General" type="String"/><column name="CREATE TABLE" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="CREATE TABLE" property="Address" value="$O$4" type="String"/><column name="CREATE TABLE" property="ColumnWidth" value="15" type="Double"/><column name="CREATE TABLE" property="NumberFormat" value="General" type="String"/><column name="CREATE VIEW" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="CREATE VIEW" property="Address" value="$P$4" type="String"/><column name="CREATE VIEW" property="ColumnWidth" value="14.43" type="Double"/><column name="CREATE VIEW" property="NumberFormat" value="General" type="String"/><column name="ALTER PROCEDURE" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="ALTER PROCEDURE" property="Address" value="$Q$4" type="String"/><column name="ALTER PROCEDURE" property="ColumnWidth" value="8.43" type="Double"/><column name="ALTER PROCEDURE" property="NumberFormat" value="General" type="String"/><column name="CREATE FUNCTION" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="CREATE FUNCTION" property="Address" value="$R$4" type="String"/><column name="CREATE FUNCTION" property="ColumnWidth" value="19.29" type="Double"/><column name="CREATE FUNCTION" property="NumberFormat" value="General" type="String"/><column name="ALTER ANY USER" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="ALTER ANY USER" property="Address" value="$S$4" type="String"/><column name="ALTER ANY USER" property="ColumnWidth" value="17.29" type="Double"/><column name="ALTER ANY USER" property="NumberFormat" value="General" type="String"/><column name="ALTER ANY ROLE" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="ALTER ANY ROLE" property="Address" value="$T$4" type="String"/><column name="ALTER ANY ROLE" property="ColumnWidth" value="17.14" type="Double"/><column name="ALTER ANY ROLE" property="NumberFormat" value="General" type="String"/><column name="CREATE ROLE" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="CREATE ROLE" property="Address" value="$U$4" type="String"/><column name="CREATE ROLE" property="ColumnWidth" value="14" type="Double"/><column name="CREATE ROLE" property="NumberFormat" value="General" type="String"/><column name="TAKE OWNERSHIP" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TAKE OWNERSHIP" property="Address" value="$V$4" type="String"/><column name="TAKE OWNERSHIP" property="ColumnWidth" value="18.57" type="Double"/><column name="TAKE OWNERSHIP" property="NumberFormat" value="General" type="String"/><column name="ALTER" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="ALTER" property="Address" value="$W$4" type="String"/><column name="ALTER" property="ColumnWidth" value="7.86" type="Double"/><column name="ALTER" property="NumberFormat" value="General" type="String"/><column name="CONTROL" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="CONTROL" property="Address" value="$X$4" type="String"/><column name="CONTROL" property="ColumnWidth" value="11" type="Double"/><column name="CONTROL" property="NumberFormat" value="General" type="String"/><column name="CONNECT" property="FormatConditions(1).ColumnsCount" value="20" type="Double"/><column name="CONNECT" property="FormatConditions(1).AppliesTo.Address" value="$E$4:$X$33" type="String"/><column name="CONNECT" property="FormatConditions(1).Type" value="1" type="Double"/><column name="CONNECT" property="FormatConditions(1).Priority" value="7" type="Double"/><column name="CONNECT" property="FormatConditions(1).StopIfTrue" value="True" type="Boolean"/><column name="CONNECT" property="FormatConditions(1).Formula1" value="=&quot;DENY+&quot;" type="String"/><column name="CONNECT" property="FormatConditions(1).Operator" value="3" type="Double"/><column name="CONNECT" property="FormatConditions(1).NumberFormat" value="General" type="String"/><column name="CONNECT" property="FormatConditions(1).Font.Bold" value="True" type="Boolean"/><column name="CONNECT" property="FormatConditions(1).Interior.Color" value="255" type="Double"/><column name="CONNECT" property="FormatConditions(1).Interior.Color" value="255" type="Double"/><column name="CONNECT" property="FormatConditions(2).ColumnsCount" value="20" type="Double"/><column name="CONNECT" property="FormatConditions(2).AppliesTo.Address" value="$E$4:$X$33" type="String"/><column name="CONNECT" property="FormatConditions(2).Type" value="1" type="Double"/><column name="CONNECT" property="FormatConditions(2).Priority" value="8" type="Double"/><column name="CONNECT" property="FormatConditions(2).StopIfTrue" value="True" type="Boolean"/><column name="CONNECT" property="FormatConditions(2).Formula1" value="=&quot;DENY&quot;" type="String"/><column name="CONNECT" property="FormatConditions(2).Operator" value="3" type="Double"/><column name="CONNECT" property="FormatConditions(2).NumberFormat" value="General" type="String"/><column name="CONNECT" property="FormatConditions(2).Interior.Color" value="255" type="Double"/><column name="CONNECT" property="FormatConditions(2).Interior.Color" value="255" type="Double"/><column name="CONNECT" property="FormatConditions(3).ColumnsCount" value="20" type="Double"/><column name="CONNECT" property="FormatConditions(3).AppliesTo.Address" value="$E$4:$X$33" type="String"/><column name="CONNECT" property="FormatConditions(3).Type" value="1" type="Double"/><column name="CONNECT" property="FormatConditions(3).Priority" value="9" type="Double"/><column name="CONNECT" property="FormatConditions(3).StopIfTrue" value="True" type="Boolean"/><column name="CONNECT" property="FormatConditions(3).Formula1" value="=&quot;GRANT+&quot;" type="String"/><column name="CONNECT" property="FormatConditions(3).Operator" value="3" type="Double"/><column name="CONNECT" property="FormatConditions(3).NumberFormat" value="General" type="String"/><column name="CONNECT" property="FormatConditions(3).Font.Bold" value="True" type="Boolean"/><column name="CONNECT" property="FormatConditions(3).Interior.Color" value="5287936" type="Double"/><column name="CONNECT" property="FormatConditions(3).Interior.Color" value="5287936" type="Double"/><column name="CONNECT" property="FormatConditions(4).ColumnsCount" value="20" type="Double"/><column name="CONNECT" property="FormatConditions(4).AppliesTo.Address" value="$E$4:$X$33" type="String"/><column name="CONNECT" property="FormatConditions(4).Type" value="1" type="Double"/><column name="CONNECT" property="FormatConditions(4).Priority" value="10" type="Double"/><column name="CONNECT" property="FormatConditions(4).StopIfTrue" value="True" type="Boolean"/><column name="CONNECT" property="FormatConditions(4).Formula1" value="=&quot;GRANT&quot;" type="String"/><column name="CONNECT" property="FormatConditions(4).Operator" value="3" type="Double"/><column name="CONNECT" property="FormatConditions(4).NumberFormat" value="General" type="String"/><column name="CONNECT" property="FormatConditions(4).Interior.Color" value="5287936" type="Double"/><column name="CONNECT" property="FormatConditions(4).Interior.Color" value="5287936" type="Double"/><column name="CONNECT" property="FormatConditions(5).ColumnsCount" value="20" type="Double"/><column name="CONNECT" property="FormatConditions(5).AppliesTo.Address" value="$E$4:$X$33" type="String"/><column name="CONNECT" property="FormatConditions(5).Type" value="9" type="Double"/><column name="CONNECT" property="FormatConditions(5).Priority" value="11" type="Double"/><column name="CONNECT" property="FormatConditions(5).StopIfTrue" value="True" type="Boolean"/><column name="CONNECT" property="FormatConditions(5).Text" value="DENY" type="String"/><column name="CONNECT" property="FormatConditions(5).TextOperator" value="0" type="Double"/><column name="CONNECT" property="FormatConditions(5).Font.Color" value="393372" type="Double"/><column name="CONNECT" property="FormatConditions(5).Font.Color" value="393372" type="Double"/><column name="CONNECT" property="FormatConditions(5).Interior.Color" value="13551615" type="Double"/><column name="CONNECT" property="FormatConditions(5).Interior.Color" value="13551615" type="Double"/><column name="CONNECT" property="FormatConditions(6).ColumnsCount" value="20" type="Double"/><column name="CONNECT" property="FormatConditions(6).AppliesTo.Address" value="$E$4:$X$33" type="String"/><column name="CONNECT" property="FormatConditions(6).Type" value="9" type="Double"/><column name="CONNECT" property="FormatConditions(6).Priority" value="12" type="Double"/><column name="CONNECT" property="FormatConditions(6).StopIfTrue" value="True" type="Boolean"/><column name="CONNECT" property="FormatConditions(6).Text" value="GRANT" type="String"/><column name="CONNECT" property="FormatConditions(6).TextOperator" value="0" type="Double"/><column name="CONNECT" property="FormatConditions(6).Font.Color" value="24832" type="Double"/><column name="CONNECT" property="FormatConditions(6).Font.Color" value="24832" type="Double"/><column name="CONNECT" property="FormatConditions(6).Interior.Color" value="13561798" type="Double"/><column name="CONNECT" property="FormatConditions(6).Interior.Color" value="13561798" type="Double"/><column name="" property="ActiveWindow.DisplayGridlines" value="False" type="Boolean"/><column name="" property="ActiveWindow.FreezePanes" value="True" type="Boolean"/><column name="" property="ActiveWindow.Split" value="True" type="Boolean"/><column name="" property="ActiveWindow.SplitRow" value="0" type="Double"/><column name="" property="ActiveWindow.SplitColumn" value="-2" type="Double"/><column name="" property="PageSetup.Orientation" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesWide" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesTall" value="1" type="Double"/></columnFormats></table>');
INSERT INTO xls.formats (TABLE_SCHEMA, TABLE_NAME, TABLE_EXCEL_FORMAT_XML) VALUES (N'xls', N'usp_principal_permissions', N'<table name="xls.usp_principal_permissions"><columnFormats><column name="" property="ListObjectName" value="principal_permissions" type="String"/><column name="" property="ShowTotals" value="False" type="Boolean"/><column name="" property="TableStyle.Name" value="TableStyleMedium15" type="String"/><column name="" property="ShowTableStyleColumnStripes" value="False" type="Boolean"/><column name="" property="ShowTableStyleFirstColumn" value="False" type="Boolean"/><column name="" property="ShowShowTableStyleLastColumn" value="False" type="Boolean"/><column name="" property="ShowTableStyleRowStripes" value="False" type="Boolean"/><column name="_RowNum" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="_RowNum" property="Address" value="$B$4" type="String"/><column name="_RowNum" property="NumberFormat" value="General" type="String"/><column name="principal_type" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="principal_type" property="Address" value="$C$4" type="String"/><column name="principal_type" property="ColumnWidth" value="15.43" type="Double"/><column name="principal_type" property="NumberFormat" value="General" type="String"/><column name="principal" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="principal" property="Address" value="$D$4" type="String"/><column name="principal" property="ColumnWidth" value="26.71" type="Double"/><column name="principal" property="NumberFormat" value="General" type="String"/><column name="name" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="name" property="Address" value="$E$4" type="String"/><column name="name" property="ColumnWidth" value="26.71" type="Double"/><column name="name" property="NumberFormat" value="General" type="String"/><column name="type" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="type" property="Address" value="$F$4" type="String"/><column name="type" property="ColumnWidth" value="13" type="Double"/><column name="type" property="NumberFormat" value="General" type="String"/><column name="has_any" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="has_any" property="Address" value="$G$4" type="String"/><column name="has_any" property="ColumnWidth" value="9.71" type="Double"/><column name="has_any" property="NumberFormat" value="General" type="String"/><column name="VIEW DEFINITION" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="VIEW DEFINITION" property="Address" value="$H$4" type="String"/><column name="VIEW DEFINITION" property="ColumnWidth" value="18.29" type="Double"/><column name="VIEW DEFINITION" property="NumberFormat" value="General" type="String"/><column name="IMPERSONATE" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="IMPERSONATE" property="Address" value="$I$4" type="String"/><column name="IMPERSONATE" property="ColumnWidth" value="15.57" type="Double"/><column name="IMPERSONATE" property="NumberFormat" value="General" type="String"/><column name="TAKE OWNERSHIP" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TAKE OWNERSHIP" property="Address" value="$J$4" type="String"/><column name="TAKE OWNERSHIP" property="ColumnWidth" value="18.57" type="Double"/><column name="TAKE OWNERSHIP" property="NumberFormat" value="General" type="String"/><column name="ALTER" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="ALTER" property="Address" value="$K$4" type="String"/><column name="ALTER" property="ColumnWidth" value="7.86" type="Double"/><column name="ALTER" property="NumberFormat" value="General" type="String"/><column name="CONTROL" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="CONTROL" property="Address" value="$L$4" type="String"/><column name="CONTROL" property="ColumnWidth" value="11" type="Double"/><column name="CONTROL" property="NumberFormat" value="General" type="String"/><column name="VIEW DEFINITION" property="FormatConditions(1).ColumnsCount" value="5" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(1).AppliesTo.Address" value="$H$4:$L$16" type="String"/><column name="VIEW DEFINITION" property="FormatConditions(1).Type" value="1" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(1).Priority" value="7" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(1).StopIfTrue" value="True" type="Boolean"/><column name="VIEW DEFINITION" property="FormatConditions(1).Formula1" value="=&quot;DENY+&quot;" type="String"/><column name="VIEW DEFINITION" property="FormatConditions(1).Operator" value="3" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(1).NumberFormat" value="General" type="String"/><column name="VIEW DEFINITION" property="FormatConditions(1).Font.Bold" value="True" type="Boolean"/><column name="VIEW DEFINITION" property="FormatConditions(1).Interior.Color" value="255" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(1).Interior.Color" value="255" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(2).ColumnsCount" value="5" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(2).AppliesTo.Address" value="$H$4:$L$16" type="String"/><column name="VIEW DEFINITION" property="FormatConditions(2).Type" value="1" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(2).Priority" value="8" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(2).StopIfTrue" value="True" type="Boolean"/><column name="VIEW DEFINITION" property="FormatConditions(2).Formula1" value="=&quot;DENY&quot;" type="String"/><column name="VIEW DEFINITION" property="FormatConditions(2).Operator" value="3" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(2).NumberFormat" value="General" type="String"/><column name="VIEW DEFINITION" property="FormatConditions(2).Interior.Color" value="255" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(2).Interior.Color" value="255" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(3).ColumnsCount" value="5" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(3).AppliesTo.Address" value="$H$4:$L$16" type="String"/><column name="VIEW DEFINITION" property="FormatConditions(3).Type" value="1" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(3).Priority" value="9" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(3).StopIfTrue" value="True" type="Boolean"/><column name="VIEW DEFINITION" property="FormatConditions(3).Formula1" value="=&quot;GRANT+&quot;" type="String"/><column name="VIEW DEFINITION" property="FormatConditions(3).Operator" value="3" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(3).NumberFormat" value="General" type="String"/><column name="VIEW DEFINITION" property="FormatConditions(3).Font.Bold" value="True" type="Boolean"/><column name="VIEW DEFINITION" property="FormatConditions(3).Interior.Color" value="5287936" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(3).Interior.Color" value="5287936" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(4).ColumnsCount" value="5" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(4).AppliesTo.Address" value="$H$4:$L$16" type="String"/><column name="VIEW DEFINITION" property="FormatConditions(4).Type" value="1" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(4).Priority" value="10" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(4).StopIfTrue" value="True" type="Boolean"/><column name="VIEW DEFINITION" property="FormatConditions(4).Formula1" value="=&quot;GRANT&quot;" type="String"/><column name="VIEW DEFINITION" property="FormatConditions(4).Operator" value="3" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(4).NumberFormat" value="General" type="String"/><column name="VIEW DEFINITION" property="FormatConditions(4).Interior.Color" value="5287936" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(4).Interior.Color" value="5287936" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(5).ColumnsCount" value="5" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(5).AppliesTo.Address" value="$H$4:$L$16" type="String"/><column name="VIEW DEFINITION" property="FormatConditions(5).Type" value="9" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(5).Priority" value="11" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(5).StopIfTrue" value="True" type="Boolean"/><column name="VIEW DEFINITION" property="FormatConditions(5).Text" value="DENY" type="String"/><column name="VIEW DEFINITION" property="FormatConditions(5).TextOperator" value="0" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(5).Font.Color" value="393372" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(5).Font.Color" value="393372" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(5).Interior.Color" value="13551615" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(5).Interior.Color" value="13551615" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(6).ColumnsCount" value="5" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(6).AppliesTo.Address" value="$H$4:$L$16" type="String"/><column name="VIEW DEFINITION" property="FormatConditions(6).Type" value="9" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(6).Priority" value="12" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(6).StopIfTrue" value="True" type="Boolean"/><column name="VIEW DEFINITION" property="FormatConditions(6).Text" value="GRANT" type="String"/><column name="VIEW DEFINITION" property="FormatConditions(6).TextOperator" value="0" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(6).Font.Color" value="24832" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(6).Font.Color" value="24832" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(6).Interior.Color" value="13561798" type="Double"/><column name="VIEW DEFINITION" property="FormatConditions(6).Interior.Color" value="13561798" type="Double"/><column name="" property="ActiveWindow.DisplayGridlines" value="False" type="Boolean"/><column name="" property="ActiveWindow.FreezePanes" value="True" type="Boolean"/><column name="" property="ActiveWindow.Split" value="True" type="Boolean"/><column name="" property="ActiveWindow.SplitRow" value="0" type="Double"/><column name="" property="ActiveWindow.SplitColumn" value="-2" type="Double"/><column name="" property="PageSetup.Orientation" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesWide" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesTall" value="1" type="Double"/></columnFormats></table>');
INSERT INTO xls.formats (TABLE_SCHEMA, TABLE_NAME, TABLE_EXCEL_FORMAT_XML) VALUES (N'xls', N'usp_object_permissions', N'<table name="xls.usp_object_permissions"><columnFormats><column name="" property="ListObjectName" value="object_permissions" type="String"/><column name="" property="ShowTotals" value="False" type="Boolean"/><column name="" property="TableStyle.Name" value="TableStyleMedium15" type="String"/><column name="" property="ShowTableStyleColumnStripes" value="False" type="Boolean"/><column name="" property="ShowTableStyleFirstColumn" value="False" type="Boolean"/><column name="" property="ShowShowTableStyleLastColumn" value="False" type="Boolean"/><column name="" property="ShowTableStyleRowStripes" value="False" type="Boolean"/><column name="_RowNum" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="_RowNum" property="Address" value="$B$4" type="String"/><column name="_RowNum" property="NumberFormat" value="General" type="String"/><column name="principal_type" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="principal_type" property="Address" value="$C$4" type="String"/><column name="principal_type" property="ColumnWidth" value="15.71" type="Double"/><column name="principal_type" property="NumberFormat" value="General" type="String"/><column name="principal" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="principal" property="Address" value="$D$4" type="String"/><column name="principal" property="ColumnWidth" value="26.71" type="Double"/><column name="principal" property="NumberFormat" value="General" type="String"/><column name="schema" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="schema" property="Address" value="$E$4" type="String"/><column name="schema" property="ColumnWidth" value="13.57" type="Double"/><column name="schema" property="NumberFormat" value="General" type="String"/><column name="name" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="name" property="Address" value="$F$4" type="String"/><column name="name" property="ColumnWidth" value="51.43" type="Double"/><column name="name" property="NumberFormat" value="General" type="String"/><column name="type" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="type" property="Address" value="$G$4" type="String"/><column name="type" property="ColumnWidth" value="15.71" type="Double"/><column name="type" property="NumberFormat" value="General" type="String"/><column name="has_any" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="has_any" property="Address" value="$H$4" type="String"/><column name="has_any" property="ColumnWidth" value="9.71" type="Double"/><column name="has_any" property="NumberFormat" value="General" type="String"/><column name="has_direct" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="has_direct" property="Address" value="$I$4" type="String"/><column name="has_direct" property="ColumnWidth" value="11.71" type="Double"/><column name="has_direct" property="NumberFormat" value="General" type="String"/><column name="SELECT" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="SELECT" property="Address" value="$J$4" type="String"/><column name="SELECT" property="ColumnWidth" value="8.57" type="Double"/><column name="SELECT" property="NumberFormat" value="General" type="String"/><column name="INSERT" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="INSERT" property="Address" value="$K$4" type="String"/><column name="INSERT" property="ColumnWidth" value="8.71" type="Double"/><column name="INSERT" property="NumberFormat" value="General" type="String"/><column name="UPDATE" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="UPDATE" property="Address" value="$L$4" type="String"/><column name="UPDATE" property="ColumnWidth" value="9.71" type="Double"/><column name="UPDATE" property="NumberFormat" value="General" type="String"/><column name="DELETE" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="DELETE" property="Address" value="$M$4" type="String"/><column name="DELETE" property="ColumnWidth" value="8.71" type="Double"/><column name="DELETE" property="NumberFormat" value="General" type="String"/><column name="EXECUTE" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="EXECUTE" property="Address" value="$N$4" type="String"/><column name="EXECUTE" property="ColumnWidth" value="10.29" type="Double"/><column name="EXECUTE" property="NumberFormat" value="General" type="String"/><column name="VIEW DEFINITION" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="VIEW DEFINITION" property="Address" value="$O$4" type="String"/><column name="VIEW DEFINITION" property="ColumnWidth" value="18.29" type="Double"/><column name="VIEW DEFINITION" property="NumberFormat" value="General" type="String"/><column name="REFERENCES" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="REFERENCES" property="Address" value="$P$4" type="String"/><column name="REFERENCES" property="ColumnWidth" value="13.43" type="Double"/><column name="REFERENCES" property="NumberFormat" value="General" type="String"/><column name="TAKE OWNERSHIP" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TAKE OWNERSHIP" property="Address" value="$Q$4" type="String"/><column name="TAKE OWNERSHIP" property="ColumnWidth" value="18.57" type="Double"/><column name="TAKE OWNERSHIP" property="NumberFormat" value="General" type="String"/><column name="ALTER" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="ALTER" property="Address" value="$R$4" type="String"/><column name="ALTER" property="ColumnWidth" value="7.86" type="Double"/><column name="ALTER" property="NumberFormat" value="General" type="String"/><column name="CONTROL" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="CONTROL" property="Address" value="$S$4" type="String"/><column name="CONTROL" property="ColumnWidth" value="11" type="Double"/><column name="CONTROL" property="NumberFormat" value="General" type="String"/><column name="SELECT" property="FormatConditions(1).ColumnsCount" value="10" type="Double"/><column name="SELECT" property="FormatConditions(1).AppliesTo.Address" value="$J$4:$S$89" type="String"/><column name="SELECT" property="FormatConditions(1).Type" value="1" type="Double"/><column name="SELECT" property="FormatConditions(1).Priority" value="7" type="Double"/><column name="SELECT" property="FormatConditions(1).StopIfTrue" value="True" type="Boolean"/><column name="SELECT" property="FormatConditions(1).Formula1" value="=&quot;DENY+&quot;" type="String"/><column name="SELECT" property="FormatConditions(1).Operator" value="3" type="Double"/><column name="SELECT" property="FormatConditions(1).NumberFormat" value="General" type="String"/><column name="SELECT" property="FormatConditions(1).Font.Bold" value="True" type="Boolean"/><column name="SELECT" property="FormatConditions(1).Interior.Color" value="255" type="Double"/><column name="SELECT" property="FormatConditions(1).Interior.Color" value="255" type="Double"/><column name="SELECT" property="FormatConditions(2).ColumnsCount" value="10" type="Double"/><column name="SELECT" property="FormatConditions(2).AppliesTo.Address" value="$J$4:$S$89" type="String"/><column name="SELECT" property="FormatConditions(2).Type" value="1" type="Double"/><column name="SELECT" property="FormatConditions(2).Priority" value="8" type="Double"/><column name="SELECT" property="FormatConditions(2).StopIfTrue" value="True" type="Boolean"/><column name="SELECT" property="FormatConditions(2).Formula1" value="=&quot;DENY&quot;" type="String"/><column name="SELECT" property="FormatConditions(2).Operator" value="3" type="Double"/><column name="SELECT" property="FormatConditions(2).NumberFormat" value="General" type="String"/><column name="SELECT" property="FormatConditions(2).Interior.Color" value="255" type="Double"/><column name="SELECT" property="FormatConditions(2).Interior.Color" value="255" type="Double"/><column name="SELECT" property="FormatConditions(3).ColumnsCount" value="10" type="Double"/><column name="SELECT" property="FormatConditions(3).AppliesTo.Address" value="$J$4:$S$89" type="String"/><column name="SELECT" property="FormatConditions(3).Type" value="1" type="Double"/><column name="SELECT" property="FormatConditions(3).Priority" value="9" type="Double"/><column name="SELECT" property="FormatConditions(3).StopIfTrue" value="True" type="Boolean"/><column name="SELECT" property="FormatConditions(3).Formula1" value="=&quot;GRANT+&quot;" type="String"/><column name="SELECT" property="FormatConditions(3).Operator" value="3" type="Double"/><column name="SELECT" property="FormatConditions(3).NumberFormat" value="General" type="String"/><column name="SELECT" property="FormatConditions(3).Font.Bold" value="True" type="Boolean"/><column name="SELECT" property="FormatConditions(3).Interior.Color" value="5287936" type="Double"/><column name="SELECT" property="FormatConditions(3).Interior.Color" value="5287936" type="Double"/><column name="SELECT" property="FormatConditions(4).ColumnsCount" value="10" type="Double"/><column name="SELECT" property="FormatConditions(4).AppliesTo.Address" value="$J$4:$S$89" type="String"/><column name="SELECT" property="FormatConditions(4).Type" value="1" type="Double"/><column name="SELECT" property="FormatConditions(4).Priority" value="10" type="Double"/><column name="SELECT" property="FormatConditions(4).StopIfTrue" value="True" type="Boolean"/><column name="SELECT" property="FormatConditions(4).Formula1" value="=&quot;GRANT&quot;" type="String"/><column name="SELECT" property="FormatConditions(4).Operator" value="3" type="Double"/><column name="SELECT" property="FormatConditions(4).NumberFormat" value="General" type="String"/><column name="SELECT" property="FormatConditions(4).Interior.Color" value="5287936" type="Double"/><column name="SELECT" property="FormatConditions(4).Interior.Color" value="5287936" type="Double"/><column name="SELECT" property="FormatConditions(5).ColumnsCount" value="10" type="Double"/><column name="SELECT" property="FormatConditions(5).AppliesTo.Address" value="$J$4:$S$89" type="String"/><column name="SELECT" property="FormatConditions(5).Type" value="9" type="Double"/><column name="SELECT" property="FormatConditions(5).Priority" value="11" type="Double"/><column name="SELECT" property="FormatConditions(5).StopIfTrue" value="True" type="Boolean"/><column name="SELECT" property="FormatConditions(5).Text" value="DENY" type="String"/><column name="SELECT" property="FormatConditions(5).TextOperator" value="0" type="Double"/><column name="SELECT" property="FormatConditions(5).Font.Color" value="393372" type="Double"/><column name="SELECT" property="FormatConditions(5).Font.Color" value="393372" type="Double"/><column name="SELECT" property="FormatConditions(5).Interior.Color" value="13551615" type="Double"/><column name="SELECT" property="FormatConditions(5).Interior.Color" value="13551615" type="Double"/><column name="SELECT" property="FormatConditions(6).ColumnsCount" value="10" type="Double"/><column name="SELECT" property="FormatConditions(6).AppliesTo.Address" value="$J$4:$S$89" type="String"/><column name="SELECT" property="FormatConditions(6).Type" value="9" type="Double"/><column name="SELECT" property="FormatConditions(6).Priority" value="12" type="Double"/><column name="SELECT" property="FormatConditions(6).StopIfTrue" value="True" type="Boolean"/><column name="SELECT" property="FormatConditions(6).Text" value="GRANT" type="String"/><column name="SELECT" property="FormatConditions(6).TextOperator" value="0" type="Double"/><column name="SELECT" property="FormatConditions(6).Font.Color" value="24832" type="Double"/><column name="SELECT" property="FormatConditions(6).Font.Color" value="24832" type="Double"/><column name="SELECT" property="FormatConditions(6).Interior.Color" value="13561798" type="Double"/><column name="SELECT" property="FormatConditions(6).Interior.Color" value="13561798" type="Double"/><column name="" property="ActiveWindow.DisplayGridlines" value="False" type="Boolean"/><column name="" property="ActiveWindow.FreezePanes" value="True" type="Boolean"/><column name="" property="ActiveWindow.Split" value="True" type="Boolean"/><column name="" property="ActiveWindow.SplitRow" value="0" type="Double"/><column name="" property="ActiveWindow.SplitColumn" value="-2" type="Double"/><column name="" property="PageSetup.Orientation" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesWide" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesTall" value="1" type="Double"/></columnFormats></table>');
INSERT INTO xls.formats (TABLE_SCHEMA, TABLE_NAME, TABLE_EXCEL_FORMAT_XML) VALUES (N'xls', N'usp_role_members', N'<table name="xls.usp_role_members"><columnFormats><column name="" property="ListObjectName" value="role_members" type="String"/><column name="" property="ShowTotals" value="False" type="Boolean"/><column name="" property="TableStyle.Name" value="TableStyleMedium15" type="String"/><column name="" property="ShowTableStyleColumnStripes" value="False" type="Boolean"/><column name="" property="ShowTableStyleFirstColumn" value="False" type="Boolean"/><column name="" property="ShowShowTableStyleLastColumn" value="False" type="Boolean"/><column name="" property="ShowTableStyleRowStripes" value="False" type="Boolean"/><column name="_RowNum" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="_RowNum" property="Address" value="$B$4" type="String"/><column name="_RowNum" property="NumberFormat" value="General" type="String"/><column name="type" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="type" property="Address" value="$C$4" type="String"/><column name="type" property="ColumnWidth" value="13" type="Double"/><column name="type" property="NumberFormat" value="General" type="String"/><column name="name" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="name" property="Address" value="$D$4" type="String"/><column name="name" property="ColumnWidth" value="26.71" type="Double"/><column name="name" property="NumberFormat" value="General" type="String"/><column name="format_column" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="format_column" property="Address" value="$E$4" type="String"/><column name="format_column" property="NumberFormat" value="General" type="String"/><column name="format_column" property="HorizontalAlignment" value="-4108" type="Double"/><column name="format_column" property="Font.Size" value="10" type="Double"/><column name="xls_admins" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="xls_admins" property="Address" value="$F$4" type="String"/><column name="xls_admins" property="ColumnWidth" value="12.43" type="Double"/><column name="xls_admins" property="NumberFormat" value="General" type="String"/><column name="xls_admins" property="HorizontalAlignment" value="-4108" type="Double"/><column name="xls_admins" property="Font.Size" value="10" type="Double"/><column name="xls_developers" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="xls_developers" property="Address" value="$G$4" type="String"/><column name="xls_developers" property="ColumnWidth" value="16.14" type="Double"/><column name="xls_developers" property="NumberFormat" value="General" type="String"/><column name="xls_developers" property="HorizontalAlignment" value="-4108" type="Double"/><column name="xls_developers" property="Font.Size" value="10" type="Double"/><column name="xls_formats" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="xls_formats" property="Address" value="$H$4" type="String"/><column name="xls_formats" property="ColumnWidth" value="12.86" type="Double"/><column name="xls_formats" property="NumberFormat" value="General" type="String"/><column name="xls_formats" property="HorizontalAlignment" value="-4108" type="Double"/><column name="xls_formats" property="Font.Size" value="10" type="Double"/><column name="xls_users" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="xls_users" property="Address" value="$I$4" type="String"/><column name="xls_users" property="ColumnWidth" value="10.71" type="Double"/><column name="xls_users" property="NumberFormat" value="General" type="String"/><column name="xls_users" property="HorizontalAlignment" value="-4108" type="Double"/><column name="xls_users" property="Font.Size" value="10" type="Double"/><column name="last_format_column" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="last_format_column" property="Address" value="$J$4" type="String"/><column name="last_format_column" property="NumberFormat" value="General" type="String"/><column name="last_format_column" property="HorizontalAlignment" value="-4108" type="Double"/><column name="format_column" property="FormatConditions(1).ColumnsCount" value="6" type="Double"/><column name="format_column" property="FormatConditions(1).AppliesTo.Address" value="$E$4:$J$27" type="String"/><column name="format_column" property="FormatConditions(1).Type" value="6" type="Double"/><column name="format_column" property="FormatConditions(1).Priority" value="13" type="Double"/><column name="format_column" property="FormatConditions(1).ShowIconOnly" value="True" type="Boolean"/><column name="format_column" property="FormatConditions(1).IconSet.ID" value="8" type="Double"/><column name="format_column" property="FormatConditions(1).IconCriteria(1).Type" value="3" type="Double"/><column name="format_column" property="FormatConditions(1).IconCriteria(1).Operator" value="7" type="Double"/><column name="format_column" property="FormatConditions(1).IconCriteria(2).Type" value="0" type="Double"/><column name="format_column" property="FormatConditions(1).IconCriteria(2).Value" value="0.5" type="Double"/><column name="format_column" property="FormatConditions(1).IconCriteria(2).Operator" value="7" type="Double"/><column name="format_column" property="FormatConditions(1).IconCriteria(3).Type" value="0" type="Double"/><column name="format_column" property="FormatConditions(1).IconCriteria(3).Value" value="1" type="Double"/><column name="format_column" property="FormatConditions(1).IconCriteria(3).Operator" value="7" type="Double"/><column name="" property="ActiveWindow.DisplayGridlines" value="False" type="Boolean"/><column name="" property="ActiveWindow.FreezePanes" value="True" type="Boolean"/><column name="" property="ActiveWindow.Split" value="True" type="Boolean"/><column name="" property="ActiveWindow.SplitRow" value="0" type="Double"/><column name="" property="ActiveWindow.SplitColumn" value="-2" type="Double"/><column name="" property="PageSetup.Orientation" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesWide" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesTall" value="1" type="Double"/></columnFormats></table>');
GO

INSERT INTO xls.workbooks (NAME, TEMPLATE, DEFINITION, TABLE_SCHEMA) VALUES (N'savetodb_permissions.xlsx', NULL,
N'database_permissions=xls.usp_database_permissions,(Default),False,$B$3,,{"Parameters":{},"ListObjectName":"database_permissions"}
principal_permissions=xls.usp_principal_permissions,(Default),False,$B$3,,{"Parameters":{"principal":null,"name":null,"has_any":1},"ListObjectName":"principal_permissions"}
object_permissions=xls.usp_object_permissions,(Default),False,$B$3,,{"Parameters":{"principal":null,"schema":null,"type":"","has_any":null,"has_direct":1},"ListObjectName":"object_permissions"}
role_members=xls.usp_role_members,(Default),False,$B$3,,{"Parameters":{},"ListObjectName":"role_members"}', N'xls');
GO

print 'SaveToDB Administrator Framework installed'
