use northwind
go
CREATE TABLE application_errors (
	error_id int IDENTITY (1, 1) PRIMARY KEY NOT NULL,
	username varchar(50) NOT NULL,
	error_message varchar(200) NOT NULL,
	error_time smalldatetime NOT NULL,
      stack_trace varchar(4000) NULL)
go
CREATE PROCEDURE usp_application_errors_save
@user_name varchar(50),
@error_message varchar(200),
@error_time smalldatetime,
@stack_trace varchar(4000)
AS
INSERT INTO application_errors
  (username, error_message, error_time, stack_trace)
  VALUES
  (@user_name, @error_message, @error_time, @stack_trace)
