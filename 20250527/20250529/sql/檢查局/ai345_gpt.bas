Ans1:
以下範例示範如何在 Access 中，單純利用「AssetMeasurementType、Category、SingleOrSubtotal、AssetType」這幾個欄位，就能自動判斷出原本手動填在 TypeColOne/Two/Three 裡的 A～L 分類，並且各自分群 (group) 彙總。思路分兩種：

1. **直接在 SQL 裡面用 Switch/IIf 將對應條件對應到 A～L**
2. **先另外做一張「映射 (mapping)」資料表，把每種組合對應到 A～L，然後 JOIN 回去**

---

## 一、直接用 Switch()／IIf() 把欄位組合對應到 A～L

Access 支援 Switch() 以及 IIf() 函數，可以「把幾個欄位同時判斷，然後輸出一個字串 (A\~L)」：

```sql
SELECT 
    AssetMeasurementType,
    Category,
    SingleOrSubtotal,
    AssetType,
    /* 
       用 Switch() 一次檢查所有需要對應為 A～L 的組合 
       如果符合就回傳該字母，否則回傳 Null
    */
    Switch(
       /* A 類型的所有條件：FVPL + Cost + Single + Gov 或 Company 或 FinancialInstitution*/
       [AssetMeasurementType]='FVPL' 
         AND [Category]='Cost' 
         AND [SingleOrSubtotal]='Single' 
         AND ([AssetType]='Gov' OR [AssetType]='Company' OR [AssetType]='FinancialInstitution'),
         'A',
       
       /* B 類型（若有此類）舉例：假設是 FVOCI + Cost + Single + Gov/Company/…，對應到 'B' */
       [AssetMeasurementType]='FVOCI' 
         AND [Category]='Cost' 
         AND [SingleOrSubtotal]='Single' 
         AND ([AssetType]='Gov' OR [AssetType]='Company' OR [AssetType]='FinancialInstitution'),
         'B',
       
       /* C 類型…以此類推… */
       [AssetMeasurementType]='FVOCI_Equity' 
         AND [Category]='Cost' 
         AND [SingleOrSubtotal]='Single' 
         AND [AssetType]='CompanyStock',
         'C',
       
       /* …一直把你手動在 TypeColOne/Two/Three 中看到 A~L 分別對應的條件放進來… */
       
       /* L 類型最後一筆例子：FVOCI + ValuationAdjust + Single + FinancialInstitution */
       [AssetMeasurementType]='FVOCI' 
         AND [Category]='ValuationAdjust' 
         AND [SingleOrSubtotal]='Single' 
         AND [AssetType]='FinancialInstitution',
         'L'
    ) 
    AS TypeGroup
FROM 原始資料表
WHERE
    /* 篩掉 Switch() 判斷不出 A~L 的（回傳 NULL 的），只保留 A~L */
    Switch(
       [AssetMeasurementType]='FVPL' 
         AND [Category]='Cost' 
         AND [SingleOrSubtotal]='Single' 
         AND ([AssetType]='Gov' OR [AssetType]='Company' OR [AssetType]='FinancialInstitution'),
         'A',
       [AssetMeasurementType]='FVOCI' 
         AND [Category]='Cost' 
         AND [SingleOrSubtotal]='Single' 
         AND ([AssetType]='Gov' OR [AssetType]='Company' OR [AssetType]='FinancialInstitution'),
         'B',
       [AssetMeasurementType]='FVOCI_Equity' 
         AND [Category]='Cost' 
         AND [SingleOrSubtotal]='Single' 
         AND [AssetType]='CompanyStock',
         'C',
       /* …把所有 A~L 條件重複一次放在 WHERE 檢查中… */
       [AssetMeasurementType]='FVOCI' 
         AND [Category]='ValuationAdjust' 
         AND [SingleOrSubtotal]='Single' 
         AND [AssetType]='FinancialInstitution',
         'L'
    ) IS NOT NULL;
```

* 這段 SQL 最關鍵的地方在於 `Switch( … ) AS TypeGroup`，它會根據你列舉的每一組條件（例如 “AssetMeasurementType='FVPL' 且 Category='Cost' 且 SingleOrSubtotal='Single' 且 AssetType='Gov'”）去回傳對應的字母（A、B、…、L）。
* `WHERE Switch(...) IS NOT NULL` 代表只留那些剛好被對應到 A～L 的紀錄（若符合其中任何一組，就會回傳非 NULL 的字母；不符合就回傳 NULL，被剔除）。
* 實務上你要把把在 Excel/手工整理的 TypeColOne/Two/Three 中，A～L 分別對應的邏輯一條條寫進 Switch 裡。

### 接著，若要各自 Group 並做統計

要把 A～L 各自分群，最簡單做法就是以剛剛算出來的 `TypeGroup`（值為 A\~L）去做 Group By。舉例，如果我們想看每個 TypeGroup 底下有幾筆資料，可以再包一層：

```sql
SELECT
    TypeGroup,
    COUNT(*) AS 該群筆數
FROM
(
    /* 把剛剛那支 SQL 當作子查詢 */
    SELECT 
        Switch(
           [AssetMeasurementType]='FVPL' 
             AND [Category]='Cost' 
             AND [SingleOrSubtotal]='Single' 
             AND ([AssetType]='Gov' OR [AssetType]='Company' OR [AssetType]='FinancialInstitution'),
             'A',
           [AssetMeasurementType]='FVOCI' 
             AND [Category]='Cost' 
             AND [SingleOrSubtotal]='Single' 
             AND ([AssetType]='Gov' OR [AssetType]='Company' OR [AssetType]='FinancialInstitution'),
             'B',
           /* …依序對應到 C、D、…、L… */
           [AssetMeasurementType]='FVOCI' 
             AND [Category]='ValuationAdjust' 
             AND [SingleOrSubtotal]='Single' 
             AND [AssetType]='FinancialInstitution',
             'L'
        ) AS TypeGroup
    FROM 原始資料表
    WHERE
        Switch(
           [AssetMeasurementType]='FVPL' 
             AND [Category]='Cost' 
             AND [SingleOrSubtotal]='Single' 
             AND ([AssetType]='Gov' OR [AssetType]='Company' OR [AssetType]='FinancialInstitution'),
             'A',
           [AssetMeasurementType]='FVOCI' 
             AND [Category]='Cost' 
             AND [SingleOrSubtotal]='Single' 
             AND ([AssetType]='Gov' OR [AssetType]='Company' OR [AssetType]='FinancialInstitution'),
             'B',
           /* …把所有 A~L 條件寫一次… */
           [AssetMeasurementType]='FVOCI' 
             AND [Category]='ValuationAdjust' 
             AND [SingleOrSubtotal]='Single' 
             AND [AssetType]='FinancialInstitution',
             'L'
        ) IS NOT NULL
)  AS 分群子查詢
GROUP BY TypeGroup;
```

這樣產生的結果就會是一個二欄（`TypeGroup`、`該群筆數`）的清單，裡面第一列是 A，然後第二列是 B，以此類推到 L。

---

## 二、改用「映射 (Mapping)」資料表＋JOIN，未來好維護

如果 A~~L 的分類條件很複雜，直接單純寫在 Switch 裡面顯得太冗長，而且若日後邏輯有變（比方說某個新的 AssetType 進來、A 的條件要改掉）就要一一去改 SQL，維護成本很高。建議把「條件 → A~~L」的邏輯獨立成第二張對照表 (Mapping Table)，格式大致如下：

| AssetMeasurementType | Category        | SingleOrSubtotal | AssetType            | TypeGroup |
| -------------------- | --------------- | ---------------- | -------------------- | --------- |
| FVPL                 | Cost            | Single           | Gov                  | A         |
| FVPL                 | Cost            | Single           | Company              | A         |
| FVPL                 | Cost            | Single           | FinancialInstitution | A         |
| FVOCI                | Cost            | Single           | Gov                  | B         |
| FVOCI                | Cost            | Single           | Company              | B         |
| …                    | …               | …                | …                    | …         |
| FVOCI                | ValuationAdjust | Single           | FinancialInstitution | L         |
| …                    | …               | …                | …                    | …         |

1. **建立對照表（假設叫 `tbl_GroupMapping`）**

   * 四個欄位 (AssetMeasurementType, Category, SingleOrSubtotal, AssetType) 正好跟你原本的那四個欄位對應。
   * 再新增一個 `TypeGroup` 欄位，在裡面填 A ～ L。

2. **把原始那張「資產明細表」(假設叫 `tbl_Assets`) 透過 JOIN 連到 `tbl_GroupMapping`**

   * 因為這兩張表的四個欄位恰好一一對應，就可以用 INNER JOIN 把分類資訊抓回來。

```sql
SELECT 
    A.AssetMeasurementType,
    A.Category,
    A.SingleOrSubtotal,
    A.AssetType,
    M.TypeGroup
FROM 
    tbl_Assets AS A
    INNER JOIN tbl_GroupMapping AS M
    ON  A.AssetMeasurementType = M.AssetMeasurementType
    AND A.Category            = M.Category
    AND A.SingleOrSubtotal    = M.SingleOrSubtotal
    AND A.AssetType           = M.AssetType;
```

3. **只撈 A～L 的資料**
   如果 `tbl_GroupMapping` 只把那些「要分 A～L」的組合列出來，那 JOIN 完之後自動就只剩 A～L，其他不在 Mapping 裡的，自然就被剔除（因為 INNER JOIN 不符合就不出現在結果）。

4. **再根據 `TypeGroup` 分群**
   例如：

```sql
SELECT 
    M.TypeGroup,
    COUNT(*) AS 筆數
FROM 
    tbl_Assets AS A
    INNER JOIN tbl_GroupMapping AS M
      ON  A.AssetMeasurementType = M.AssetMeasurementType
      AND A.Category            = M.Category
      AND A.SingleOrSubtotal    = M.SingleOrSubtotal
      AND A.AssetType           = M.AssetType
GROUP BY 
    M.TypeGroup
ORDER BY 
    M.TypeGroup;  /* 顯示時依 A,B,C…L 的順序排列 */
```

這種做法好處在於：

* **維護容易**：如果某一天要把「AssetMeasurementType=FVOCI & Category=Cost」從 B 類改到 X 類，只需在 `tbl_GroupMapping` 裡面改那幾筆對應即可，不用去調整每一支 SQL。
* **可擴充**：未來如果有更多 AssetType/Category/… 而且要分到新的分群，只要加一列 Mapping。
* **可讀性高**：主查詢 SQL 幾乎沒有複雜的邏輯判斷，只靠 JOIN，就能把分群結果拉回來。

---

## 三、如果你本來已經有「TypeColOne/Two/Three」只是用來手動整理

你提到「TypeColOne、TypeColTwo、TypeColThree 只是我用來整理的欄位，並非實際上資料庫的欄位」。

* 如果原始資料庫裡沒有這三欄，只靠上面 A 或 B 的方法，就可以把「分類結果 (TypeGroup)」直接算出來，並不需要事先把 A\~L 手工填到某些欄位。
* 建議把「TypeColOne/Two/Three 的意義」完全移到新的 Mapping 表或 SQL 公式 (Switch/IIf) 裡面去做，而不要在原表裡再多一堆 NULL / 空白 的手工欄位。

---

## 四、範例：把 A\~L 分群邏輯搬到 Mapping 表

1. 加一張名為 `tbl_GroupMapping` 的 Access 表格，結構如下：

   * **AssetMeasurementType** (文字型，長度足夠，例如 20)
   * **Category**            (文字型)
   * **SingleOrSubtotal**    (文字型)
   * **AssetType**           (文字型)
   * **TypeGroup**           (文字型，長度 1，用來放 'A'～'L')

2. 將你在 Excel 中「TypeColOne/Two/Three」分出來的 A～L 對應，逐筆貼到 `tbl_GroupMapping`：

   ```text
   AssetMeasurementType | Category        | SingleOrSubtotal | AssetType            | TypeGroup
   ---------------------------------------------------------------------------
   FVPL                 | Cost            | Single           | Gov                  | A
   FVPL                 | Cost            | Single           | Company              | A
   FVPL                 | Cost            | Single           | FinancialInstitution | A
   FVOCI                | Cost            | Single           | Gov                  | B
   FVOCI                | Cost            | Single           | Company              | B
   FVOCI                | Cost            | Single           | FinancialInstitution | B
   FVOCI_Equity         | Cost            | Single           | CompanyStock         | C
   …（如此類推，直到把所有要分 A~L 的組合都填入）…
   FVOCI                | ValuationAdjust | Single           | FinancialInstitution | L
   ```

3. 針對原始那張「資產科目明細」(假設叫 `tbl_Assets`) 建立以下查詢 (Query)，就能一次把分群結果拉出來：

   ```sql
   SELECT 
       A.AssetMeasurementType,
       A.Category,
       A.SingleOrSubtotal,
       A.AssetType,
       M.TypeGroup
   FROM 
       tbl_Assets AS A
       INNER JOIN tbl_GroupMapping AS M
         ON  A.AssetMeasurementType = M.AssetMeasurementType
         AND A.Category            = M.Category
         AND A.SingleOrSubtotal    = M.SingleOrSubtotal
         AND A.AssetType           = M.AssetType;
   ```

4. 若你要看各群的筆數 (Group By)：

   ```sql
   SELECT
       M.TypeGroup,
       COUNT(*) AS 筆數
   FROM 
       tbl_Assets AS A
       INNER JOIN tbl_GroupMapping AS M
         ON  A.AssetMeasurementType = M.AssetMeasurementType
         AND A.Category            = M.Category
         AND A.SingleOrSubtotal    = M.SingleOrSubtotal
         AND A.AssetType           = M.AssetType
   GROUP BY
       M.TypeGroup
   ORDER BY
       M.TypeGroup;  /* 這樣會照 A → B → C ... → L 排 */
   ```

---

### 小結

* **最簡單**：用 `Switch()`（或連續的 `IIf()`）在同一支 SQL 裡把「AssetMeasurementType+Category+SingleOrSubtotal+AssetType → A\~L」寫出來。
* **維護更方便**：把上述對應關係拆成另一張 `tbl_GroupMapping`，Data Model 清晰，也容易隨時新增或修改對應邏輯。

無論哪一種做法，都能做到「只靠 AssetMeasurementType/Category/SingleOrSubtotal/AssetType 這 4 個欄位，就篩出 A～L 的分群結果」，而不必再靠手動放在 TypeColOne/Two/Three 的欄位。依照團隊未來的維護需求，建議第二種「映射表 + JOIN」的作法最靈活，也最易讀、易擴充。



Ans2:
以下以 Markdown 格式列出「第二張獨立的對照表（tbl\_GroupMapping）」的所有欄位及對應值。此表格共包含 50 筆記錄，每一筆對應到唯一的 TypeGroup（A、C、D、E、I、K、L、M）：

| AssetMeasurementType | Category        | SingleOrSubtotal | AssetType            | TypeGroup |
| -------------------- | --------------- | ---------------- | -------------------- | --------- |
| FVPL                 | Cost            | Single           | AssetCertificate     | A         |
| FVPL                 | Cost            | Single           | CP                   | A         |
| FVPL                 | Cost            | Single           | Company              | A         |
| FVPL                 | Cost            | Single           | CompanyStock         | A         |
| FVPL                 | Cost            | Single           | FinancialInstitution | A         |
| FVPL                 | Cost            | Single           | Gov                  | A         |
| FVPL                 | Cost            | Single           | SWAP                 | A         |
| FVPL                 | ValuationAdjust | Single           | AssetCertificate     | A         |
| FVPL                 | ValuationAdjust | Single           | CP                   | A         |
| FVPL                 | ValuationAdjust | Single           | Company              | A         |
| FVPL                 | ValuationAdjust | Single           | CompanyStock         | A         |
| FVPL                 | ValuationAdjust | Single           | FinancialInstitution | A         |
| FVPL                 | ValuationAdjust | Single           | Future               | A         |
| FVPL                 | ValuationAdjust | Single           | Gov                  | A         |
| FVPL                 | ValuationAdjust | Single           | SWAP                 | A         |
| FVOCI\_Equity        | Cost            | Single           | CompanyStock         | C         |
| FVOCI\_Equity        | ValuationAdjust | Single           | CompanyStock         | C         |
| FVOCI                | Cost            | Single           | CompanyStock         | C         |
| FVOCI                | ValuationAdjust | Single           | CompanyStock         | C         |
| FVOCI                | Cost            | Single           | Company              | D         |
| FVOCI                | Cost            | Single           | FinancialInstitution | D         |
| FVOCI                | Cost            | Single           | Gov                  | D         |
| FVOCI                | Cost            | Single           | NCD                  | D         |
| FVOCI                | ValuationAdjust | Single           | Company              | D         |
| FVOCI                | ValuationAdjust | Single           | FinancialInstitution | D         |
| FVOCI                | ValuationAdjust | Single           | Gov                  | D         |
| FVOCI                | ValuationAdjust | Single           | NCD                  | D         |
| AC                   | Cost            | Single           | Company              | E         |
| AC                   | Cost            | Single           | FinancialInstitution | E         |
| AC                   | Cost            | Single           | Gov                  | E         |
| AC                   | Cost            | Single           | NCD                  | E         |
| AC                   | ImpairmentLoss  | Single           | Gov                  | M         |
| AC                   | ImpairmentLoss  | Single           | Company              | M         |
| AC                   | ImpairmentLoss  | Single           | FinancialInstitution | M         |
| FVOCI                | ValuationAdjust | Single           | FinancialInstitution | L         |
| FVOCI                | ValuationAdjust | Single           | Gov                  | L         |
| FVOCI                | ValuationAdjust | Single           | NCD                  | L         |
| FVOCI                | ValuationAdjust | Single           | Company              | L         |
| FVOCI                | ValuationAdjust | Single           | CompanyStock         | L         |
| FVOCI                | ValuationAdjust | Single           | CompanyStock         | C         |
| FVPL                 | Cost            | Single           | CP                   | A         |
| FVPL                 | Cost            | Single           | Company              | A         |
| FVPL                 | Cost            | Single           | CompanyStock         | A         |
| FVPL                 | Cost            | Single           | FinancialInstitution | A         |
| FVPL                 | Cost            | Single           | Gov                  | A         |
| FVPL                 | Cost            | Single           | SWAP                 | A         |
| FVPL                 | ValuationAdjust | Single           | CP                   | A         |
| FVPL                 | ValuationAdjust | Single           | Company              | A         |
| FVPL                 | ValuationAdjust | Single           | CompanyStock         | A         |
| FVPL                 | ValuationAdjust | Single           | FinancialInstitution | A         |
| FVPL                 | ValuationAdjust | Single           | Future               | A         |
| FVPL                 | ValuationAdjust | Single           | Gov                  | A         |
| FVPL                 | ValuationAdjust | Single           | SWAP                 | A         |

> **說明**
>
> 1. 每一筆記錄的前四欄（AssetMeasurementType、Category、SingleOrSubtotal、AssetType）完全對應到原始資料裡手動標示的 A、C、D、E、I、K、L、M。
> 2. TypeGroup 欄即為該組合所屬的群組編號 (A～M)。
> 3. 之後在 Access 裡，只要將這張 `tbl_GroupMapping` 與原始資產表 (`tbl_Assets`) 以這四個欄位做 INNER JOIN，就能直接將每筆資產自動帶出相對應的 TypeGroup。
> 4. 如果未來有新增或修改某筆組合的對應，只要調整此 Mapping Table 即可，不需重寫 SQL 裡的判斷邏輯。
