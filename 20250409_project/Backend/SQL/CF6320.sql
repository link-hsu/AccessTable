Query FB5A_OBU_CF6320_LIST
SELECT *
From OBU_CF6320
WHERE OBU_CF6320.DataMonthString = "2024/11";
      
      
PARAMETERS DataMonthParam TEXT;
SELECT *
From OBU_CF6320
WHERE OBU_CF6320.DataMonthString = [DataMonthParam]; 
      
      
      
      
      
      
 
