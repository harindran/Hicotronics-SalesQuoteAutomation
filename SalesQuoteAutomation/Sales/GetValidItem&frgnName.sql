create Procedure "GetValidItemCode"(In ItemCode varchar(1000))
as
begin
SELECT  distinct CASE  
WHEN EXISTS(SELECT 1 FROM OITM WHERE  "ItemCode"=:ItemCode)
        THEN (SELECT "ItemCode" from OITM WHERE  "ItemCode"=:ItemCode)                
END "ItemCode"  from OITM;
        end;
        
        Create Procedure "GetValidFrgnCode"(In FrgnName varchar(1000))as
begin
SELECT  distinct CASE  
WHEN EXISTS(SELECT 1 FROM OITM WHERE  "FrgnName"=:FrgnName)
        THEN (SELECT "ItemCode" from OITM WHERE  "FrgnName"=:FrgnName)      
        else
        '13723'          
END "ItemCode"  from OITM;
        end;
        
