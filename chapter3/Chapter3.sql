use northwind
go
create PROCEDURE usp_region_getall 
AS
SELECT  *
FROM	Region;
go
--Retrieve a single record from the Region table:
create PROCEDURE usp_region_getone
@id int 
AS
SELECT  *
FROM	Region
WHERE 	RegionID = @id;
go
--Delete a single record from the Region table:
create PROCEDURE usp_region_delete
@id int
AS
DELETE
FROM	Region
WHERE	RegionID = @id;
go
--Save a record to the Region table (this includes both inserts and updates):
create PROCEDURE usp_region_save
@id int,
@region varchar(50),
@new_id int OUTPUT
AS

IF @id = 0
  BEGIN
    SET @id = (SELECT MAX(RegionID) 
	FROM Region) + 1

    INSERT INTO Region (RegionID, RegionDescription)
      VALUES (@id, @region)
  END
ELSE
  UPDATE Region
  SET	 RegionDescription = @region
  WHERE	 RegionID = @id

SET @new_id = @id;
