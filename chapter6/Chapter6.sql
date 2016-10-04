Use Northwind
CREATE TABLE menus (
   menu_id int identity(1,1) primary key,
   menu_under_id int null foreign key references menus(menu_id),
   menu_order int not null,
   menu_caption varchar(30) not null,
   menu_shortcut varchar(5),
   enabled int default(1) check(enabled in (0,1)),
   checked int default(0) check(checked in (0,1)))
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
null,1,'&File',null)
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
null,2,'&Edit',null)
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
null,3,'&Maintenance',null)
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
null,4,'&Window',null)
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
null,5,'&Help',null)
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
1,1,'E&xit',null)
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
2,1,'Cu&t','CtrlX')
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
2,2,'&Copy','CtrlC')
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
2,3,'&Paste','CtrlV')
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
2,4,'-',null)
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
2,5,'Select Al&l','CtrlA')
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
2,6,'&Find…','CtrlF')
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
2,7,'Find &Next','F3')
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
3,1,'&Regions',null)
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
3,2,'&Territories',null)
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
3,3,'&Employees',null)
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
4,1,'&Cascade',null)
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
4,2,'&Tile',null)
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
18,1,'&Horizontal',null)
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
18,2,'&Vertical',null)
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
5,1,'&Report Errors',null)
go
INSERT INTO menus (
menu_under_id, menu_order, menu_caption, menu_shortcut) VALUES (
5,2,'&About',null)
go
CREATE PROCEDURE get_menu_structure
AS
SELECT *
FROM	 menus
ORDER BY menu_under_id, menu_order