CREATE Procedure "GetValidItemCode"(In ItemCode varchar(1000))
as
begin
SELECT  distinct CASE  WHEN EXISTS(SELECT 1 FROM OITM WHERE  "ItemCode"=:ItemCode)
        THEN (SELECT "ItemCode" from OITM WHERE  "ItemCode"=:ItemCode)
        ELSE '13723'  END "ItemCode"  from OITM;
        end;


		CREATE Procedure "GetValidMPN_MakeCode"(In MPN varchar(1000),In Make varchar(1000))as
begin
SELECT  distinct CASE  
WHEN EXISTS(SELECT 1 FROM OITM WHERE  "U_OrderPN"=:MPN and "U_Make"=:Make)
        THEN (SELECT Top 1 "ItemCode" from OITM WHERE  "U_OrderPN"=:MPN and "U_Make"=:Make)          
        else
        '13723'          
END "ItemCode"  from OITM;
        end;