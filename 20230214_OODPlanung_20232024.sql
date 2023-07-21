USE [OOD]
GO
IF OBJECT_ID(N'[dbo].[oodplanung2023]') IS NOT NULL
	DROP TABLE [dbo].[oodplanung2023]
GO
CREATE TABLE [dbo].[oodplanung2023](
	[datefrom] [datetime] NOT NULL,
	[dateuntil] [datetime] NOT NULL,
	[weeknofrom]  AS (datepart(iso_week,[datefrom])),
	[weeknountil]  AS (datepart(iso_week,[dateuntil])),
	[weekdayfrom]  AS (datepart(weekday,[datefrom])),
	[weekdayuntil]  AS (datepart(weekday,[dateuntil])),
	[who] [nvarchar](5) NULL
)
GO
DECLARE @datefrom datetime = {ts'2023-03-01 13:00:00.000'}, @dateuntil datetime = {ts'2023-03-02 12:00:00.000'}
INSERT INTO [dbo].[oodplanung2023]([datefrom],[dateuntil],[who]) VALUES(@datefrom, @dateuntil,N'KEL')
WHILE (SELECT MAX([dateuntil]) FROM [dbo].[oodplanung2023]) < {d'2024-03-01' }
	BEGIN
		SET @datefrom = DATEADD(hh,1,(SELECT MAX([dateuntil]) FROM [dbo].[oodplanung2023]))
		SET @dateuntil = DATEADD(hh,-1,@datefrom +1)
		INSERT INTO [dbo].[oodplanung2023]([datefrom],[dateuntil]) VALUES(@datefrom, @dateuntil)
	END
GO
--Wochenenden 
UPDATE [dbo].[oodplanung2023] SET who = NULL
UPDATE [dbo].[oodplanung2023] SET who = N'we' WHERE weekdayfrom IN (7,1) OR weekdayuntil in (1)
--SELECT * FROM [OOD].[dbo].[oodplanung2023] WHERE weekdayfrom = 6 and weekdayuntil = 7
UPDATE [dbo].[oodplanung2023] SET [dateuntil] = dateuntil +2 WHERE weekdayfrom = 6 and weekdayuntil = 7
DELETE FROM [dbo].[oodplanung2023] WHERE who = N'we'




--Feiertage
--Ostern 2023-04-07 - 2023-04-10 
DELETE FROM [OOD].[dbo].[oodplanung2023] WHERE datefrom IN( {ts'2023-04-07 13:00:00.000'}, {ts'2023-04-10 13:00:00.000'})
--Dienst Nachmittag von 2023-04-06 geht bis 2023-04-11 
UPDATE [OOD].[dbo].[oodplanung2023] SET [dateuntil] = {ts'2023-04-11 12:00:00.000'}  WHERE datefrom = {ts'2023-04-06 13:00:00.000'}
--Auffahrt 2023-05-18
DELETE FROM [OOD].[dbo].[oodplanung2023] WHERE datefrom IN( {ts'2023-05-18 13:00:00.000'})
UPDATE [OOD].[dbo].[oodplanung2023] SET [dateuntil] = {ts'2023-05-19 12:00:00.000'}  WHERE datefrom = {ts'2023-05-17 13:00:00.000'}
--Auffahrt 2023-05-29
DELETE FROM [OOD].[dbo].[oodplanung2023] WHERE datefrom IN( {ts'2023-05-29 13:00:00.000'})
UPDATE [OOD].[dbo].[oodplanung2023] SET [dateuntil] = {ts'2023-05-30 12:00:00.000'}  WHERE datefrom = {ts'2023-05-26 13:00:00.000'}
--Nationalfeiertag 2023-08-01
DELETE FROM [OOD].[dbo].[oodplanung2023] WHERE datefrom IN( {ts'2023-08-01 13:00:00.000'})
UPDATE [OOD].[dbo].[oodplanung2023] SET [dateuntil] = {ts'2023-08-02 12:00:00.000'}  WHERE datefrom = {ts'2023-07-31 13:00:00.000'}
--Weihnachten 2023-12-25 - 2023-12-26
DELETE FROM [OOD].[dbo].[oodplanung2023] WHERE datefrom IN( {ts'2023-12-25 13:00:00.000'}, {ts'2023-12-26 13:00:00.000'})
UPDATE [OOD].[dbo].[oodplanung2023] SET [dateuntil] = {ts'2023-12-27 12:00:00.000'}  WHERE datefrom = {ts'2023-12-22 13:00:00.000'}
--Silvester 2024-01-01 - 2024-01-01
DELETE FROM [OOD].[dbo].[oodplanung2023] WHERE datefrom IN( {ts'2024-01-01 13:00:00.000'}, {ts'2024-01-02 13:00:00.000'})
UPDATE [OOD].[dbo].[oodplanung2023] SET [dateuntil] = {ts'2024-01-03 12:00:00.000'}  WHERE datefrom = {ts'2023-12-29 13:00:00.000'}


GO

IF OBJECT_ID(N'[dbo].[oodplanung2023_Participants]') IS NOT NULL
	DROP TABLE [dbo].[oodplanung2023_Participants]
GO
CREATE TABLE [dbo].[oodplanung2023_Participants]
	(
	[who] [nvarchar](5) NOT NULL
	,[mail] [nvarchar](260) NOT NULL
	,[pensum] smallint not null
	)
INSERT INTO [dbo].[oodplanung2023_Participants]([who],[mail],[pensum]) VALUES(N'GWY',N'guido.wymann@bedag.ch', 90)
INSERT INTO [dbo].[oodplanung2023_Participants]([who],[mail],[pensum]) VALUES(N'BRU',N'bruno.rodrigues@bedag.ch',100)
INSERT INTO [dbo].[oodplanung2023_Participants]([who],[mail],[pensum]) VALUES(N'KEL',N'kevin.luginbuehl@bedag.ch',100)
INSERT INTO [dbo].[oodplanung2023_Participants]([who],[mail],[pensum]) VALUES(N'MCI',N'michelle.vonkaenel@bedag.ch',80)
INSERT INTO [dbo].[oodplanung2023_Participants]([who],[mail],[pensum]) VALUES(N'SCR',N'roman.schneider@bedag.ch',100)
INSERT INTO [dbo].[oodplanung2023_Participants]([who],[mail],[pensum]) VALUES(N'JSP',N'juerg.spring@bedag.ch',100)
INSERT INTO [dbo].[oodplanung2023_Participants]([who],[mail],[pensum]) VALUES(N'MMAR',N'mario.markotic@bedag.ch',100)
INSERT INTO [dbo].[oodplanung2023_Participants]([who],[mail],[pensum]) VALUES(N'ASTE',N'anton.steiner@bedag.ch',90)
INSERT INTO [dbo].[oodplanung2023_Participants]([who],[mail],[pensum]) VALUES(N'LOH',N'loris.held@bedag.ch',100)
GO
IF OBJECT_ID(N'[dbo].[oodplanung2023_absence]') IS NOT NULL
	DROP TABLE [dbo].[oodplanung2023_absence]
GO
CREATE TABLE [dbo].[oodplanung2023_absence] (
	[datefrom] date null,
	[dateuntil] date null,
	[weekday] smallint null,
	[weekno] smallint null,
	[cntable] bit not null,
	[who] nvarchar(5) not null)
GO
--standardfreitag
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES (null, null, 6, null, 1, N'GWY')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES (null, null, 2, null, 1, N'ASTE')

--jede abwesenheit inkl. Tag davor
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-03-24'}, {d'2023-04-02'}, null, null, 0, N'GWY')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-06-09'}, {d'2023-06-25'}, null, null, 0, N'GWY')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-08-18'}, {d'2023-09-03'}, null, null, 0, N'GWY')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-09-22'}, {d'2023-10-01'}, null, null, 0, N'GWY')

INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-03-10'}, {d'2023-03-13'}, null, null, 0, N'MCI')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-05-05'}, {d'2023-05-21'}, null, null, 0, N'MCI')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-06-02'}, {d'2023-06-11'}, null, null, 0, N'MCI')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-06-27'}, {d'2023-07-16'}, null, null, 0, N'MCI')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-07-28'}, {d'2023-08-01'}, null, null, 0, N'MCI')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-08-30'}, {d'2023-09-03'}, null, null, 0, N'MCI')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-09-20'}, {d'2023-09-24'}, null, null, 0, N'MCI')

INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-05-19'}, {d'2023-06-04'}, null, null, 0, N'KEL')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-06-07'}, {d'2023-06-11'}, null, null, 0, N'KEL')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-09-22'}, {d'2023-10-08'}, null, null, 0, N'KEL')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-12-22'}, {d'2024-01-08'}, null, null, 0, N'KEL')

INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-07-07'}, {d'2023-08-06'}, null, null, 0, N'JSP')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-09-29'}, {d'2023-10-15'}, null, null, 0, N'JSP')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-12-22'}, {d'2024-01-02'}, null, null, 0, N'JSP')

INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-05-17'}, {d'2023-06-04'}, null, null, 0, N'ASTE')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-08-18'}, {d'2023-09-10'}, null, null, 0, N'ASTE')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-08-18'}, {d'2023-09-10'}, null, null, 0, N'ASTE')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-12-22'}, {d'2024-01-07'}, null, null, 0, N'ASTE')

INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-04-06'}, {d'2023-04-16'}, null, null, 0, N'SCR')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-06-29'}, {d'2023-07-02'}, null, null, 0, N'SCR')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-07-14'}, {d'2023-08-06'}, null, null, 0, N'SCR')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-08-09'}, {d'2023-08-13'}, null, null, 0, N'SCR')

INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-02-10'}, {d'2023-03-12'}, null, null, 1, N'MMAR')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-07-07'}, {d'2023-07-23'}, null, null, 0, N'MMAR')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-07-07'}, {d'2023-07-23'}, null, null, 0, N'MMAR')

INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-04-28'}, {d'2023-05-14'}, null, null, 0, N'BRU')
INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-06-09'}, {d'2023-07-16'}, null, null, 0, N'BRU')

INSERT INTO [dbo].[oodplanung2023_absence]([datefrom],[dateuntil],[weekday],[weekno],[cntable],[who])VALUES ({d'2023-02-10'}, {d'2023-06-04'}, null, null, 1, N'LOH')


UPDATE [dbo].[oodplanung2023] SET who = N'MCI' WHERE [datefrom] = {ts'2023-03-01 13:00:00.000'}
GO


IF OBJECT_ID(N'[dbo].[oodplanung2023_dist]') IS NOT NULL
DROP TABLE [dbo].[oodplanung2023_dist]
GO
IF OBJECT_ID(N'[dbo].[oodplanung2023_whonot]') IS NOT NULL
DROP TABLE [dbo].[oodplanung2023_whonot]
GO
SELECT
	a.*, 
	whonot = [b].[who]
INTO
	[dbo].[oodplanung2023_whonot]
FROM
	[dbo].[oodplanung2023] [a]
	LEFT JOIN [dbo].[oodplanung2023_absence] [b]
		ON CAST([a].[datefrom] as date) = [b].[datefrom]
		OR CAST([a].[datefrom] as date) = [b].[dateuntil]
		OR CAST([a].[datefrom] as date) BETWEEN [b].[datefrom] AND [b].[dateuntil]
		OR CAST([a].[dateuntil] as date) = [b].[datefrom]
		OR CAST([a].[dateuntil] as date) = [b].[dateuntil]
		OR CAST([a].[dateuntil] as date) BETWEEN [b].[datefrom] AND [b].[dateuntil]
		OR [a].[weekdayfrom] = [b].[weekday]
		OR [a].[weekdayuntil] = [b].[weekday]
		OR ([a].[weeknofrom] = [b].[weekno] AND DATEPART(YYYY, [a].[datefrom]) = 2023)
		OR ([a].[weeknountil] = [b].[weekno] AND DATEPART(YYYY, [a].[datefrom]) = 2023)
		OR ([a].[weeknofrom] = [b].[weekno] AND DATEPART(YYYY, [a].[datefrom]) = 2024)
		OR ([a].[weeknountil] = [b].[weekno] AND DATEPART(YYYY, [a].[datefrom]) = 2024)
GO
WITH [ctea] AS (
SELECT * FROM [dbo].[oodplanung2023_whonot]

), [cteb] AS (
SELECT
	*,
	[ops] = (SELECT Count(DISTINCT [datefrom]) FROM  [ctea] ),
	[nposops] = (SELECT Count(DISTINCT [datefrom]) FROM  [ctea] WHERE [whonot] = [a].[who])
FROM 
	[dbo].[oodplanung2023_Participants] [a]
), [ctec] AS (
SELECT
	*,
	[posops] = [ops] - [nposops],
	[fairpart] = ([ops] - [nposops])*[pensum]/(select sum(pensum) FROM [cteb])
FROM 
	[cteb] 
) , [cted] as (
SELECT *, 
	[miss] = ([ops] - (SELECT SUM(fairpart) from [ctec])) * [pensum] /(select sum(pensum) FROM [ctec])+1
from [ctec]
)

select 
	[who],
	[pensum],
	[fairpart] = [fairpart] + [miss]
INTO [dbo].[oodplanung2023_dist]
from 
	cted
GO
WHILE EXISTS ( SELECT * FROM [dbo].[oodplanung2023] WHERE  (DATEPART(YYYY, [datefrom]) = 2023 OR DATEPART(YYYY, [datefrom]) = 2024) and who IS NULL)
BEGIN
WITH [ctea] AS (
	SELECT 
		a.*
		, WHOlast = (SELECT [who] FROM [dbo].[oodplanung2023] WHERE [datefrom] = (SELECT MAX(datefrom) from [dbo].[oodplanung2023] WHERE [datefrom] < [a].[datefrom]) )
		,wholastweek = b.who
		,[c].[whonot]
	FROM 
		[dbo].[oodplanung2023] [a]
		LEFT JOIN [dbo].[oodplanung2023] [b]
			ON DATEPART(YYYY, [a].[datefrom]) = DATEPART(YYYY, [b].[datefrom])
			AND [a].[weeknofrom] = [b].[weeknofrom]
		LEFT JOIN [dbo].[oodplanung2023_whonot] [c]
			ON DATEPART(YYYY, [a].[datefrom]) = DATEPART(YYYY, [c].[datefrom])
			AND [a].[datefrom] = [c].[datefrom]
	WHERE 
		[a].[who] IS NULL 
		--AND DATEPART(yyyy,[a].[datefrom])=2022 
		AND [a].[datefrom] = (SELECT Min([datefrom]) from [dbo].[oodplanung2023] WHERE who is null)

), [cteb] as (
SELECT TOP 1
	[a].who
	,[datefrom] = (SELECT DISTINCT [datefrom] FROM [ctea])
FROM 
	[dbo].[oodplanung2023_dist] [a]
	LEFT JOIN (SELECT wholast from [ctea] union all SELECT wholastweek from [ctea] union all SELECT whonot from [ctea]) [b]
		ON [a].[who] = [b].[wholast]
WHERE 
	[b].WHOlast IS NULL
ORDER by
	[fairpart] - (select count(*) from [dbo].[oodplanung2023] WHERE [who] = [a].[who]) desc
)
--SELECT * FROM cteb

UPDATE [dbo].[oodplanung2023] SET who = ISNULL((SELECT who from [cteb]), N'manu')  WHERE [datefrom] = (SELECT distinct [datefrom] from [ctea])


END
GO
SELECT * FROM [dbo].[oodplanung2023] WHERE who = N'manu'
UPDATE [dbo].[oodplanung2023] SET who = N'BRU' WHERE [datefrom] = {ts'2023-03-10 13:00:00.000'}
UPDATE [dbo].[oodplanung2023] SET who = N'MMAR' WHERE [datefrom] = {ts'2023-05-05 13:00:00.000'}
UPDATE [dbo].[oodplanung2023] SET who = N'SCR' WHERE [datefrom] = {ts'2023-05-12 13:00:00.000'}
UPDATE [dbo].[oodplanung2023] SET who = N'LOH' WHERE [datefrom] = {ts'2023-07-14 13:00:00.000'}
GO
SELECT * FROM [dbo].[oodplanung2023] 
SELECT * FROM [dbo].[oodplanung2023]  WHERE weeknofrom = 28
SELECT * FROM [dbo].[oodplanung2023_whonot]
SELECT * FROM [dbo].[oodplanung2023_dist]
--SELECT * FROM [dbo].[oodplanung2023]

SELECT 
	[a].*, 
	[b].[mail],
	[crtOODnachmittag] = N'crtoodmeeting -startstring "' + CONVERT(nvarchar(19), [a].[datefrom], 121) + N'" -endstring "' + CONVERT(nvarchar(19), DATEADD(hh,4,[a].[datefrom]), 121) + N'" -t1uid "' + [b].[mail] + N'" -subj "OOD Nachmittag ' + [a].[who] + N'"',
	[crtOODvormittag] = N'crtoodmeeting -startstring "' + CONVERT(nvarchar(19), DATEADD(hh,-4,[a].[dateuntil]), 121) + N'" -endstring "' + CONVERT(nvarchar(19), [a].[dateuntil], 121) + N'" -t1uid "' + [b].[mail] + N'" -subj "OOD Vormittag ' + [a].[who] + N'"'
FROM 
	[dbo].[oodplanung2023] [a]
	INNER JOIN [dbo].[oodplanung2023_Participants] [b]
		ON [a].[who] = [b].[who]


