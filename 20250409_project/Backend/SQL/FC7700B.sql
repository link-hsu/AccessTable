Query FB5A_OBU_FC7700B_LIST

SELECT *
From OBU_FC7700B
WHERE OBU_FC7700B.DataMonthString = "2024/11";
      
      
PARAMETERS DataMonthParam TEXT;
SELECT *
From OBU_FC7700B
WHERE OBU_FC7700B.DataMonthString = [DataMonthParam];  
      
      
      
      
      
      
 
