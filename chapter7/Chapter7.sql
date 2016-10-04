USE northwind
go
CREATE PROCEDURE usp_territory_delete
@id nvarchar(20)
AS
DELETE
FROM	Territories
WHERE	TerritoryID = @id
go
CREATE PROCEDURE usp_territory_getall 
AS
SELECT  *
FROM	Territories
go
CREATE PROCEDURE usp_territory_getone
@id nvarchar(20) 
AS
SELECT  *
FROM	Territories
WHERE 	TerritoryID = @id
go
CREATE PROCEDURE usp_territory_save
@id nvarchar(20),
@territory nchar(50),
@region_id int,
@new bit
AS
IF @new = 1
   INSERT INTO Territories (TerritoryID, TerritoryDescription, RegionID)
      VALUES (@id, @territory, @region_id)
ELSE
   UPDATE  Territories
   SET	   TerritoryDescription = @territory,
	   RegionID = @region_id
   WHERE   TerritoryID = @id
