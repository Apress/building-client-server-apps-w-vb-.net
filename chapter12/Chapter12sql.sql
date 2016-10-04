USE Northwind
GO
CREATE TABLE UserList
(username VARCHAR(20) PRIMARY KEY,
 userpass VARCHAR(100) NOT NULL)
GO
CREATE PROCEDURE usp_AddNewUser
@uid varchar(20),
@pwd varchar(100)
AS
INSERT INTO UserList (username, userpass)
  VALUES (@uid, @pwd)
GO
CREATE PROCEDURE usp_ValidateUser
@uid varchar(20),
@pwd varchar(100),
@valid int OUTPUT
AS
DECLARE
@count int

SET @count = (SELECT count(*) 
	FROM 	UserList
	WHERE	username = @uid
	AND	userpass = @pwd)

IF @count = 1
   SET @valid = 1
ELSE
   SET @valid = 0

GO
exec sp_addlogin N'nwAccount', 'password', 'Northwind'
GO
--exec sp_grantdbaccess N'nwAccount', 'Northwind'

-- Jeff this is the line that needed to change, the second argument
-- of the sp_grantdbaccess refers to the name that the account is called, not the
-- database.  
exec sp_grantdbaccess N'nwAccount', 'nwAccount'  -- 
GO
GRANT  EXECUTE  ON [dbo].[usp_AddNewUser]  TO [nwAccount]
GO
GRANT  EXECUTE  ON [dbo].[usp_ValidateUser]  TO [nwAccount]
GO
GRANT  EXECUTE  ON [dbo].[usp_application_errors_save]  TO [nwAccount]
GO
GRANT  EXECUTE  ON [dbo].[usp_employee_delete]  TO [nwAccount]
GO
GRANT  EXECUTE  ON [dbo].[usp_employee_getall]  TO [nwAccount]
GO
GRANT  EXECUTE  ON [dbo].[usp_employee_getone]  TO [nwAccount]
GO
GRANT  EXECUTE  ON [dbo].[usp_employee_save]  TO [nwAccount]
GO
GRANT  EXECUTE  ON [dbo].[usp_employee_territory_delete]  TO [nwAccount]
GO
GRANT  EXECUTE  ON [dbo].[usp_employee_territory_insert]  TO [nwAccount]
GO
GRANT  EXECUTE  ON [dbo].[usp_region_delete]  TO [nwAccount]
GO
GRANT  EXECUTE  ON [dbo].[usp_region_getall]  TO [nwAccount]
GO
GRANT  EXECUTE  ON [dbo].[usp_region_getone]  TO [nwAccount]
GO
GRANT  EXECUTE  ON [dbo].[usp_region_insert]  TO [nwAccount]
GO
GRANT  EXECUTE  ON [dbo].[usp_region_save]  TO [nwAccount]
GO
GRANT  EXECUTE  ON [dbo].[usp_region_update]  TO [nwAccount]
GO
GRANT  EXECUTE  ON [dbo].[usp_territory_delete]  TO [nwAccount]
GO
GRANT  EXECUTE  ON [dbo].[usp_territory_getall]  TO [nwAccount]
GO
GRANT  EXECUTE  ON [dbo].[usp_territory_getone]  TO [nwAccount]
GO
GRANT  EXECUTE  ON [dbo].[usp_territory_save]  TO [nwAccount]
GO
GRANT  EXECUTE  ON [dbo].[get_menu_structure]  TO [nwAccount]
GO