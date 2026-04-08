# Data description for Ontario_PWQMN_2024.xlsx

## Number of sheets: 
2

## Sheet name: Data

### Column Headings
'Year', 'Collection Site', 'Analyte', 'Collected', 'Results', 'Units', 'Result Call', 'Detection Limit', 'Value Qualifier'

### First 10 Rows
|    |   Year |   Collection Site | Analyte           | Collected           | Results   | Units         | Result Call   |   Detection Limit | Value Qualifier   |
|---:|-------:|------------------:|:------------------|:--------------------|:----------|:--------------|:--------------|------------------:|:------------------|
|  0 |   2024 |       13000300102 | Potassium         | 08/14/2024          | 0.54      | mg/L          | Detected      |             0.02  | nan               |
|  1 |   2024 |       13000300102 | Molybdenum        | 08/14/2024          | <2        | ug/L          | BDL           |             2     | <                 |
|  2 |   2024 |       13001100202 | Bismuth           | 08/14/2024          | <5        | ug/L          | BDL           |             5     | <                 |
|  3 |   2024 |       13001100202 | Chromium          | 08/14/2024          | <1        | ug/L          | BDL           |             1     | <                 |
|  4 |   2024 |       13001101002 | Copper            | 08/14/2024          | 1         | ug/L          | Detected      |             0.5   | nan               |
|  5 |   2024 |       13001101002 | Zinc              | 08/14/2024          | <2        | ug/L          | BDL           |             2     | <                 |
|  6 |   2024 |        8012305202 | Alkalinity        | 2024-09-05 00:00:00 | 221.6     | mg/L as CaCO3 | Detected      |             1     | nan               |
|  7 |   2024 |        8012304802 | Conductivity      | 2024-09-05 00:00:00 | 499       | uS/cm         | Detected      |             2     | nan               |
|  8 |   2024 |        8012304802 | pH                | 2024-09-05 00:00:00 | 8.31      | nan           | Detected      |           nan     | nan               |
|  9 |   2024 |        8012305302 | Nitrogen; nitrite | 2024-09-05 00:00:00 | 0.003     | mg/L          | Detected      |             0.001 | nan               |

### Bottom 10 Rows
|        |   Year |   Collection Site | Analyte   | Collected           |   Results | Units   | Result Call   |   Detection Limit |   Value Qualifier |
|-------:|-------:|------------------:|:----------|:--------------------|----------:|:--------|:--------------|------------------:|------------------:|
| 112382 |   2024 |        3007703402 | Manganese | 09/24/2024          |      3890 | ug/L    | Detected      |               0.5 |               nan |
| 112383 |   2024 |        4001308302 | Iron      | 07/16/2024          |      2680 | ug/L    | Detected      |               3   |               nan |
| 112384 |   2024 |        4001303302 | Iron      | 07/17/2024          |      3190 | ug/L    | Detected      |               3   |               nan |
| 112385 |   2024 |        4001308202 | Iron      | 07/17/2024          |      2850 | ug/L    | Detected      |               6   |               nan |
| 112386 |   2024 |        3005703102 | Iron      | 07/24/2024          |      2670 | ug/L    | Detected      |               3   |               nan |
| 112387 |   2024 |        3005703102 | Aluminum  | 05/23/2024          |      2020 | ug/L    | Detected      |               2   |               nan |
| 112388 |   2024 |       13000800102 | Manganese | 2024-12-06 00:00:00 |      2350 | ug/L    | Detected      |               0.5 |               nan |
| 112389 |   2024 |       13000800102 | Aluminum  | 2024-12-06 00:00:00 |      3580 | ug/L    | Detected      |               2   |               nan |
| 112390 |   2024 |        4001302902 | Iron      | 04/16/2024          |       410 | ug/L    | Detected      |               3   |               nan |
| 112391 |   2024 |        4001304102 | Iron      | 07/15/2024          |      1310 | ug/L    | Detected      |               3   |               nan |

### Describe (include='all')
|        |   Year |   Collection Site | Analyte   | Collected   | Results   | Units   | Result Call   |   Detection Limit | Value Qualifier   |
|:-------|-------:|------------------:|:----------|:------------|:----------|:--------|:--------------|------------------:|:------------------|
| count  | 112392 |  112392           | 112392    | 112392      | 112392    | 108723  | 112392        |      108506       | 30119             |
| unique |    nan |     nan           | 55        | 168         | 5229      | 8       | 2             |         nan       | 1                 |
| top    |    nan |     nan           | Chloride  | 04/22/2024  | <1        | ug/L    | Detected      |         nan       | <                 |
| freq   |    nan |     nan           | 3712      | 2360        | 5685      | 50511   | 82273         |         nan       | 30119             |
| mean   |   2024 |       1.01678e+10 | nan       | nan         | nan       | nan     | nan           |           1.50042 | nan               |
| std    |      0 |       5.79876e+09 | nan       | nan         | nan       | nan     | nan           |           3.50207 | nan               |
| min    |   2024 |       1.0104e+09  | nan       | nan         | nan       | nan     | nan           |           0.001   | nan               |
| 25%    |   2024 |       6.0024e+09  | nan       | nan         | nan       | nan     | nan           |           0.05    | nan               |
| 50%    |   2024 |       8.0123e+09  | nan       | nan         | nan       | nan     | nan           |           0.5     | nan               |
| 75%    |   2024 |       1.70021e+10 | nan       | nan         | nan       | nan     | nan           |           2       | nan               |
| max    |   2024 |       1.90064e+10 | nan       | nan         | nan       | nan     | nan           |         180       | nan               |

### Info (verbose=True)
```
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 112392 entries, 0 to 112391
Data columns (total 9 columns):
 #   Column           Non-Null Count   Dtype  
---  ------           --------------   -----  
 0   Year             112392 non-null  int64  
 1   Collection Site  112392 non-null  int64  
 2   Analyte          112392 non-null  object 
 3   Collected        112392 non-null  object 
 4   Results          112392 non-null  object 
 5   Units            108723 non-null  object 
 6   Result Call      112392 non-null  object 
 7   Detection Limit  108506 non-null  float64
 8   Value Qualifier  30119 non-null   object 
dtypes: float64(1), int64(2), object(6)
memory usage: 7.7+ MB

```

### Shape
(112392, 9)

### Unique values in text columns (non-numeric only, up to 100 values)
#### Column: Analyte
'Potassium', 'Molybdenum', 'Bismuth', 'Chromium', 'Copper', 'Zinc', 'Alkalinity', 'Conductivity', 'pH', 'Nitrogen; nitrite', 'Nitrate', 'Phosphorus; total', 'Silicon; reactive silicate', 'Carbon; dissolved inorganic', 'Lead', 'Barium', 'Sodium', 'Uranium', 'Nickel', 'Nitrogen; ammonia+ammonium', 'Nitrogen; total', 'E. coli count per 100 mL', 'Iron', 'Vanadium', 'Strontium', 'Carbon; dissolved organic', 'Chloride', 'Solids; suspended', 'Lithium', 'Phosphorus; phosphate', 'Cobalt', 'Aluminum', 'Tin', 'Cadmium', 'Calcium', 'Silver', 'Nitrogen; nitrate+nitrite', 'Magnesium', 'Titanium', 'Zirconium', 'Hardness', 'Beryllium', 'Manganese', 'Sulphate', 'Selenium', 'Thallium', 'Boron', 'Antimony', 'Arsenic', 'Phenol', 'Mercury', 'Solids; dissolved', 'Solids; total', 'Sample Turbidity', 'Oxygen demand; biochemical'

#### Column: Collected
'08/14/2024', '2024-09-05 00:00:00', '05/16/2024', '2024-11-04 00:00:00', '06/18/2024', '05/22/2024', '08/13/2024', '05/28/2024', '2024-06-03 00:00:00', '2024-07-10 00:00:00', '10/17/2024', '06/17/2024', '09/24/2024', '07/16/2024', '2024-08-05 00:00:00', '07/18/2024', '06/24/2024', '06/25/2024', '10/22/2024', '2024-03-04 00:00:00', '08/20/2024', '2024-04-06 00:00:00', '2024-05-06 00:00:00', '09/16/2024', '07/23/2024', '07/15/2024', '2024-02-04 00:00:00', '08/27/2024', '06/19/2024', '2024-04-11 00:00:00', '2024-02-07 00:00:00', '10/21/2024', '10/15/2024', '09/17/2024', '07/17/2024', '08/19/2024', '08/26/2024', '2024-11-11 00:00:00', '05/13/2024', '09/23/2024', '2024-06-08 00:00:00', '11/13/2024', '05/14/2024', '11/25/2024', '11/19/2024', '08/21/2024', '2024-07-08 00:00:00', '2024-10-07 00:00:00', '2024-12-06 00:00:00', '05/21/2024', '2024-09-12 00:00:00', '2024-01-05 00:00:00', '06/13/2024', '10/29/2024', '03/18/2024', '08/15/2024', '2024-02-05 00:00:00', '2024-09-07 00:00:00', '2024-12-02 00:00:00', '04/22/2024', '04/23/2024', '2024-03-09 00:00:00', '04/15/2024', '09/18/2024', '2024-08-04 00:00:00', '05/27/2024', '02/13/2024', '05/23/2024', '2024-04-10 00:00:00', '02/28/2024', '03/25/2024', '10/16/2024', '2024-12-08 00:00:00', '02/27/2024', '10/23/2024', '2024-10-12 00:00:00', '07/24/2024', '2024-11-09 00:00:00', '2024-10-09 00:00:00', '10/31/2024', '2024-06-05 00:00:00', '2024-02-12 00:00:00', '2024-11-06 00:00:00', '03/26/2024', '11/20/2024', '2024-11-03 00:00:00', '2024-12-11 00:00:00', '2024-04-09 00:00:00', '2024-08-08 00:00:00', '2024-09-04 00:00:00', '2024-09-09 00:00:00', '08/28/2024', '2024-03-07 00:00:00', '2024-12-04 00:00:00', '2024-08-01 00:00:00', '11/26/2024', '10/30/2024', '2024-09-10 00:00:00', '2024-12-03 00:00:00', '2024-05-11 00:00:00'

#### Column: Results
'<2', '<5', '<1', '<7', '<3', '<0.5', '<9', '<4', '<14', '<0.9', '<25', '<45', '<15', '<0.1', '<1.0', '<0.02', '<18', '<0.003', '<10', '<0.2', '<0.04', '<35', '<6', '<0.001', '<27', '<0.3', '<21', '<0.05', '<0.63', '<5.0', '<0.50', '<1.5', '<30', '<20', '<0.70', '<70', '<50', '<40', '<100', '<140', '<90', '<60', '<180', '<0.4'

#### Column: Units
'mg/L', 'ug/L', 'mg/L as CaCO3', 'uS/cm', 'MPN / 100mL', 'ng/L', 'NTU', 'mg/L as O2', '<NaN>'

#### Column: Result Call
'Detected', 'BDL'

#### Column: Value Qualifier
'<', '<NaN>'

## Sheet name: Stations

### Column Headings
'STATION', 'NAME', 'LATITUDE', 'LONGITUDE', 'STATUS', 'FIRST_YEAR', 'LAST_YEAR'

### First 10 Rows
|    |    STATION | NAME                |   LATITUDE |   LONGITUDE | STATUS   |   FIRST_YEAR |   LAST_YEAR |
|---:|-----------:|:--------------------|-----------:|------------:|:---------|-------------:|------------:|
|  0 | 3007703502 | Uxbirdge Brook      |    44.1378 |    -79.1126 | A        |         2002 |        2021 |
|  1 | 3007703602 | Hawkestone Creek    |    44.4964 |    -79.4668 | A        |         2002 |        2021 |
|  2 | 3007703802 | Black River         |    44.3047 |    -79.3604 | A        |         2002 |        2021 |
|  3 | 3007703902 | Holland River       |    44.0948 |    -79.49   | A        |         2002 |        2021 |
|  4 | 3007704002 | Pefferlaw Brook     |    44.3167 |    -79.2014 | A        |         2002 |        2021 |
|  5 | 3007704102 | Beaverton River     |    44.4303 |    -79.1551 | A        |         2002 |        2021 |
|  6 | 3008500102 | Musquash River      |    45.0221 |    -79.7773 | I        |         1966 |        2002 |
|  7 | 3008500202 | Rosseau Lake Outlet |    45.1201 |    -79.5773 | I        |         1966 |        1991 |
|  8 | 3008500302 | Muskoka Lake Outlet |    45.0125 |    -79.6145 | I        |         1966 |        1995 |
|  9 | 3008500402 | Muskoka River South |    45.0024 |    -79.3021 | I        |         1965 |        1995 |

### Bottom 10 Rows
|      |     STATION | NAME                       |   LATITUDE |   LONGITUDE | STATUS   |   FIRST_YEAR |   LAST_YEAR |
|-----:|------------:|:---------------------------|-----------:|------------:|:---------|-------------:|------------:|
| 2188 | 19006404302 | Cochrane Landfill Drainage |    49.0749 |    -81.026  | I        |         1988 |        1989 |
| 2189 | 19006404402 | Cochrane Landfill Drainage |    49.0757 |    -81.0248 | I        |         1988 |        1989 |
| 2190 | 19006404502 | Porcupine River            |    48.4824 |    -81.2223 | A        |         1991 |        2021 |
| 2191 | 19006404602 | North Driftwood Creek      |    48.5377 |    -80.742  | I        |         1991 |        1996 |
| 2192 | 19006404702 | Driftwood River            |    48.5378 |    -80.687  | I        |         1991 |        1992 |
| 2193 | 19006404802 | Groundhog River            |    48.2028 |    -82.172  | I        |         2003 |        2005 |
| 2194 | 17002153802 | Trent River                |    44.1477 |    -77.5819 | A        |         2022 |           0 |
| 2195 | 16018415602 | Irvine Creek               |    43.702  |    -80.4453 | A        |         2022 |           0 |
| 2196 | 16012401302 | Venison Creek              |    42.6534 |    -80.5485 | A        |         2002 |        2021 |
| 2197 | 16018410402 | Irvine Creek               |    43.6954 |    -80.4478 | I        |         1980 |        2020 |

### Describe (include='all')
|        |        STATION | NAME         |   LATITUDE |   LONGITUDE | STATUS   |   FIRST_YEAR |   LAST_YEAR |
|:-------|---------------:|:-------------|-----------:|------------:|:---------|-------------:|------------:|
| count  | 2198           | 2198         | 2198       |  2198       | 2198     |    2198      |   2198      |
| unique |  nan           | 842          |  nan       |   nan       | 2        |     nan      |    nan      |
| top    |  nan           | Ottawa River |  nan       |   nan       | I        |     nan      |    nan      |
| freq   |  nan           | 103          |  nan       |   nan       | 1760     |     nan      |    nan      |
| mean   |    1.06727e+10 | nan          |   44.8948  |   -80.2971  | nan      |    1977.61   |   1991.97   |
| std    |    5.99731e+09 | nan          |    1.90161 |     3.23081 | nan      |      12.7569 |     62.5634 |
| min    |    1e+09       | nan          |   42.0117  |   -94.5924  | nan      |    1964      |      0      |
| 25%    |    6.0076e+09  | nan          |   43.6764  |   -81.2889  | nan      |    1968      |   1979      |
| 50%    |    9.0009e+09  | nan          |   44.2977  |   -79.9493  | nan      |    1973      |   1995      |
| 75%    |    1.70021e+10 | nan          |   45.8326  |   -78.961   | nan      |    1983      |   2005      |
| max    |    1.90064e+10 | nan          |   54.6542  |   -74.0944  | nan      |    2022      |   2021      |

### Info (verbose=True)
```
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 2198 entries, 0 to 2197
Data columns (total 7 columns):
 #   Column      Non-Null Count  Dtype  
---  ------      --------------  -----  
 0   STATION     2198 non-null   int64  
 1   NAME        2198 non-null   object 
 2   LATITUDE    2198 non-null   float64
 3   LONGITUDE   2198 non-null   float64
 4   STATUS      2198 non-null   object 
 5   FIRST_YEAR  2198 non-null   int64  
 6   LAST_YEAR   2198 non-null   int64  
dtypes: float64(2), int64(3), object(2)
memory usage: 120.3+ KB

```

### Shape
(2198, 7)

### Unique values in text columns (non-numeric only, up to 100 values)
#### Column: NAME
'Uxbirdge Brook', 'Hawkestone Creek', 'Black River', 'Holland River', 'Pefferlaw Brook', 'Beaverton River', 'Musquash River', 'Rosseau Lake Outlet', 'Muskoka Lake Outlet', 'Muskoka River South', 'Muskoka River North', 'Mary Lake Outlet', 'Fairy Lake Outlet', 'Lake Vernon Outlet', 'Lake Of Bays Outlet', 'Indian River', 'Muskoka River North Branch', 'Lake Of Bays', 'Lake Muskoka', 'Sydenham River', 'Spey River', 'Telfer Creek', 'Keefer Creek', 'Waterton Creek', 'Pigeon River', 'Bar River', 'Echo River', 'Pottawatomi River', 'Uxbridge Brook', 'Maskinonge-Jersey River', 'Maskinonge River', 'Medora Creek', 'Pigeon Creek', 'Hawk Rock River', 'Lake Rosseau', 'Dee River', 'Skeleton Creek', 'Wier Creek', 'Rosseau River', 'Shadow River', 'Lake Joseph', 'Lake Vernon', 'East River', 'Muskoka River', 'Parrot Creek', 'Moon River', 'Oastler Lake Outlet', 'Boyne River', 'Otter Lake', 'Mc Curry Lake Outlet', 'Seguin River', 'Sequin River', 'Shawanaga River', 'Orchard Creek', 'Bighead River', 'Beaver River', 'Mountain Stream West Branch', 'Mountain Stream East Branch', 'Silver Creek', 'Black Ash Creek', 'Pretty River', 'Batteaux River', 'Pine Creek', 'Pine River', 'Nottawasaga River', 'Lamont Creek', 'Beeton Creek', 'Boyne River Trib', 'Sheldon Creek', 'Mad River', 'Willow Creek', 'Innisfil Creek', 'Mcintyre Creek', 'Coates Creek', 'Bear Creek', 'Black Creek', 'Marl Creek', 'Copeland Creek', 'Wye River', 'Hog Creek', 'Sturgeon River', 'Coldwater River', 'North River', 'Drainage Canal', 'Aurora Creek', 'Canal Lake Outlet', 'Severn River', 'Schomberg River', 'Maskinonge Jersey River', 'Mount Albert Creek', 'Lake Simcoe Outlet', 'Bogart Creek', 'Lake Couchiching Outlet', 'Lovers Creek', 'Lake Superior', 'Lake Superior (Nipigon Bay)', 'Lake Superior(Jackfish Bay)', 'Lake Superior (Jackfish Bay)', 'Lake Superior (Jackfish Creek)', 'Montreal River'

#### Column: STATUS
'A', 'I'

