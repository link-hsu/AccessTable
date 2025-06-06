1.更新匯率及餘額C

-- =============================


你可以在同一個 `SELECT` 裡，利用 Access 的條件式聚合（`IIF`）來分別計算「只加總大於 0 的金額」和「只加總小於 0 的金額」。大致上把原本的


SUM(ab.Amount) AS SubtotalAmount


改成


SUM(IIF(ab.Amount > 0, ab.Amount, 0))  AS CreditBalance,   -- 貸餘（Positive）
SUM(IIF(ab.Amount < 0, ab.Amount, 0))  AS DebitBalance     -- 借餘（Negative）


完整範例如下：


PARAMETERS DataMonthParam TEXT;
SELECT
    AccountCodeMap.AssetMeasurementSubType
      & "_" & AccountCodeMap.Category      AS MeasurementCategory,
    SUM(IIF(ab.Amount > 0, ab.Amount, 0))  AS CreditBalance,   -- 貸餘
    SUM(IIF(ab.Amount < 0, ab.Amount, 0))  AS DebitBalance     -- 借餘
FROM AccountCodeMap
INNER JOIN
    (
        SELECT 
            AccountBalance.AccountCode, 
            AccountBalance.Amount
        FROM AccountBalance
        WHERE 
            AccountBalance.DataMonthString = [DataMonthParam]
            AND AccountBalance.BalanceType = '餘額C'
    ) AS ab
    ON AccountCodeMap.AccountCode = ab.AccountCode
WHERE
    AccountCodeMap.GroupFlag       IN ('台幣債','股權','受益憑證','CP','Derivative','Bill')
    AND AccountCodeMap.Category    IN ('Cost','ValuationAdjust','ImpairmentLoss','ImpairmentAllowance')
    AND AccountCodeMap.SingleOrSubtotal = 'Single'
GROUP BY
    AccountCodeMap.AssetMeasurementType,
    AccountCodeMap.AssetMeasurementSubType,
    AccountCodeMap.Category;


**說明：**

* `IIF(ab.Amount > 0, ab.Amount, 0)`：
  只有在 `ab.Amount > 0` 時才回傳該值，否則回傳 `0`，最後 `SUM` 起來就是「所有正數的合計」，也就是你的「貸餘」。

* `IIF(ab.Amount < 0, ab.Amount, 0)`：
  只有在 `ab.Amount < 0` 時才回傳該值，否則回傳 `0`，最後 `SUM` 起來就是「所有負數的合計」，也就是「借餘」。

這樣一張查詢裡就會同時回傳每個分類的「貸餘」與「借餘」兩個欄位。若你想要把「借餘」顯示成正數，可以再包一層 `ABS(...)`：

```sql
SUM(IIF(ab.Amount < 0, ABS(ab.Amount), 0)) AS DebitBalance
```

— 以上即可滿足你的需求。




-- =============================


ai345
Ⅰ類A
1100	㈠有價證券及衍生性金融資產			=F9+F10+F13+F14

***
to A
1110	 1.透過損益按公允價值衡量之金融資產	 
                =SUM(金融商品類!D3:D39)+
                SUM(金融商品類!D96:D106)
            其中SUM(金融商品類!D3:D39) 評價借方
1.透過損益按公允價值衡量之金融資產

成本
120050101	強制FVPL金融資產-公債-中央政府(我國)              
120050103	強制FVPL金融資產-公債-地方政府(我國)              
120050105	強制FVPL金融資產-公債-中央政府(外國)              
120050107	強制FVPL金融資產-公債-地方政府(外國)              
120050121	強制FVPL金融資產-普通公司債(公營)                 
120050123	強制FVPL金融資產-普通公司債(民營)                 
120050125	強制FVPL金融資產-普通公司債(公營)(外國)           
120050127	強制FVPL金融資產-普通公司債(民營)(外國)           
120050147	#N/A
120050301	強制FVPL金融資產-普通股-上市公司                  
120050302	強制FVPL金融資產-普通股-上櫃公司                  
120050311	強制FVPL金融資產-特別股-上市                      
120057307	強制FVPL金融資產-衍生性ＳＷＡＰ-ASS               
找不到 120057701	#N/A
120050501	強制FVPL金融資產-受益憑證                         
120050903	強制FVPL金融資產-商業本票                         
評價

120070101	強制FVPL金融資產評價調整-公債-中央(我國)          
120070103	強制FVPL金融資產評價調整-公債-地方(我國)          
120070105	強制FVPL金融資產評價調整-公債-中央-外國           
120070121	強制FVPL金融資產評價調整-普通公司債(公營)         
120070123	強制FVPL金融資產評價調整-普通公司債(民營)         
120070125	強制FVPL金融資產評價調整-普通公司債(公營)(外國)   
120070127	強制FVPL金融資產評價調整-普通公司債(民營)(外國)   
120070147	#N/A
120070301	強制FVPL金融資產評價調整-上市公司                 
120070302	強制FVPL金融資產評價調整-上櫃公司                 
120070311	強制FVPL金融資產評價調整-特別股-上市              
120077101	強制FVPL金融資產評價調整-遠匯(評價)               
	
120077307	強制FVPL金融資產評價調整-ＳＷＡＰ-ASS             
找不到 120077701	#N/A
120079001	強制FVPL金融資產評價調整-貸方評價調整             
120070501	強制FVPL金融資產評價調整-受益憑證                 
120070903	強制FVPL金融資產評價調整-商業本票                 

            其中SUM(金融商品類!D96:D106) 借方
(四)衍生性金融商品		
評價
120079011	強制FVPL金融資產評價調整-CVA--遠匯
120079031	強制FVPL金融資產評價調整-CVA--SWAP
120079071	強制FVPL金融資產評價調整-CVA--選擇權

***
to B
1140	2.透過其他綜合損益按公允價值衡量之金融資產			
                =SUM(F11:F12)

***
to C                
1141	A.透過其他綜合損益按公允價值衡量之權益工具			
                =SUM(金融商品類!D43:D44,金融商品類!D55:D56)                
2.透過其他綜合損益按公允價值衡量之金融資產	借方
成本	
121010301	FVOCI權益工具-普通股-上市公司                     
121019901	FVOCI權益工具-其他                                
評價
121030301	FVOCI權益工具評價調整-普通股-上市                 
121039901	FVOCI權益工具評價調整-其他                        


***
to D
1142	B.透過其他綜合損益按公允價值衡量之債務工具			
                =SUM(金融商品類!D45:D53,金融商品類!D57:D65)
借方

成本

121110101	FVOCI債務工具-公債-中央政府(我國)                 
121110103	FVOCI債務工具-公債-地方政府(我國)                 
121110105	FVOCI債務工具-公債-中央政府(外國)                 
121110121	FVOCI債務工具-普通公司債（公營）                  
121110123	FVOCI債務工具-普通公司債（民營）                  
121110125	FVOCI債務工具-普通公司債(公營)(外國)              
121110127	FVOCI債務工具-普通公司債(民營)(外國)              
121110147	FVOCI債務工具-金融債券-海外                       
121110911	FVOCI債務工具-央行NCD                             

評價
121130101	FVOCI債務工具評價調整-公債-中央政府               
121130103	FVOCI債務工具評價調整-公債-地方政府               
121130105	FVOCI債務工具評價調整-公債-中央政府(外國)         
121130121	FVOCI債務工具評價調整-普通公司債（公營)           
121130123	FVOCI債務工具評價調整-普通公司債（民營)           
121130125	FVOCI債務工具評價調整-普通公司債(公營)(外國)      
121130127	FVOCI債務工具評價調整-普通公司債(民營)(外國)      
121130147	FVOCI債務工具評價調整-金融債券-海外               
121130911	FVOCI債務工具評價調整-央行NCD                     



***
to E
1150	    3.按攤銷後成本衡量之債務工具投資			
                =SUM(金融商品類!D69:D82)

借方
3.按攤銷後成本衡量之債務工具投資	
成本
122010101	AC債務工具投資-公債-中央政府(我國)                
122010103	AC債務工具投資-公債-地方政府(我國)                
122010105	AC債務工具投資-公債-中央政府(外國)                
122010121	AC債務工具投資-普通公司債(公營)                   
122010123	AC債務工具投資-普通公司債(民營)                   
122010125	AC債務工具投資-普通公司債(公營)(外國)             
122010127	AC債務工具投資-普通公司債(民營)(外國)             
122010147	AC債務工具投資-金融債券-海外                      
122010911	AC債務工具投資-央行NCD                            
	
	
累計減損	
1220301	累積減損-AC債務工具投資-債券                      
1220309	累積減損-AC債務工具投資-票券                      


***
to F
1190	    4.其他金融資產			
                =SUM(金融商品類!D92:D93)
無

***
to G
1200	㈡採用權益法之投資			
                =SUM(金融商品類!D87:D89)
借方
(二)採用權益法之投資	
成本	
15001	採用權益法之投資成本                              
評價	
15003	加（減）：採用權益法認列之投資權益調整            


1300	㈢其他投資			
1450	㈣避險之金融資產			
1500	㈤放款	
1510	    1.逾期放款			



***

V類E


***
to H
1100	㈠有價證券及衍生性金融資產
                    =J9+J10+J13+J14	=K9+K10+K13+K14

***  
to I                  
1110	 1.透過損益按公允價值衡量之金融資產	 
                    =ABS(SUM(金融商品類!E20:E39))+
                    ABS(SUM(金融商品類!E99:E106))	
            SUM(金融商品類!E20:E39) 貸方
評價	
120070101	強制FVPL金融資產評價調整-公債-中央(我國)          
120070103	強制FVPL金融資產評價調整-公債-地方(我國)          
120070105	強制FVPL金融資產評價調整-公債-中央-外國           
120070121	強制FVPL金融資產評價調整-普通公司債(公營)         
120070123	強制FVPL金融資產評價調整-普通公司債(民營)         
120070125	強制FVPL金融資產評價調整-普通公司債(公營)(外國)   
120070127	強制FVPL金融資產評價調整-普通公司債(民營)(外國)   
120070147	#N/A
120070301	強制FVPL金融資產評價調整-上市公司                 
120070302	強制FVPL金融資產評價調整-上櫃公司                 
120070311	強制FVPL金融資產評價調整-特別股-上市              
120077101	強制FVPL金融資產評價調整-遠匯(評價)               
		
120077307	強制FVPL金融資產評價調整-ＳＷＡＰ-ASS             
120077701	#N/A
120079001	強制FVPL金融資產評價調整-貸方評價調整             
120070501	強制FVPL金融資產評價調整-受益憑證                 
120070903	強制FVPL金融資產評價調整-商業本票                 

            SUM(金融商品類!E99:E106) 貸方

(四)衍生性金融商品	
	
評價	
	
120079011	強制FVPL金融資產評價調整-CVA--遠匯
120079031	強制FVPL金融資產評價調整-CVA--SWAP
120079071	強制FVPL金融資產評價調整-CVA--選擇權
                    =J9

***
to J
1140	2.透過其他綜合損益按公允價值衡量之金融資產			
                    =SUM(J11:J12)	
                    =J10


***  
to K                  
1141	A.透過其他綜合損益按公允價值衡量之權益工具			
                    =ABS(SUM(金融商品類!E55:E56))	
貸方
2.透過其他綜合損益按公允價值衡量之金融資產
評價	
121030301	FVOCI權益工具評價調整-普通股-上市                 
121039901	FVOCI權益工具評價調整-其他                        
                    =J11                    

***   
L                 
1142	B.透過其他綜合損益按公允價值衡量之債務工具			
                    =ABS(SUM(金融商品類!E57:E65))	
2.透過其他綜合損益按公允價值衡量之金融資產
評價	
121030301	FVOCI權益工具評價調整-普通股-上市                 
121039901	FVOCI權益工具評價調整-其他                        
121130101	FVOCI債務工具評價調整-公債-中央政府               
121130103	FVOCI債務工具評價調整-公債-地方政府               
121130105	FVOCI債務工具評價調整-公債-中央政府(外國)         
121130121	FVOCI債務工具評價調整-普通公司債（公營)           
121130123	FVOCI債務工具評價調整-普通公司債（民營)           
121130125	FVOCI債務工具評價調整-普通公司債(公營)(外國)      
121130127	FVOCI債務工具評價調整-普通公司債(民營)(外國)      
121130147	FVOCI債務工具評價調整-金融債券-海外               
121130911	FVOCI債務工具評價調整-央行NCD                     
                    =J12

*** 
M                   
1150	    3.按攤銷後成本衡量之債務工具投資			
                    =ABS(SUM(金融商品類!E81:E81))	
                    
3.按攤銷後成本衡量之債務工具投資
累計減損	
1220301	累積減損-AC債務工具投資-債券                      

                    =J13

***  
N                  
1190	    4.其他金融資產			


***
O
1200	㈡採用權益法之投資			


*** 
P               
1300	㈢其他投資			

***
Q
1450	㈣避險之金融資產			

***
R
1500	㈤放款	
1510	    1.逾期放款			




***
S
1700	    7.其他				=SUM(Q37:Q227)			=28999996-158904+SUM(M37:M227)-28999996+158904	=SUM(G37:J37)	=ROUND(G37*0.02+H37*0.1+I37*0.5+J37,0)
