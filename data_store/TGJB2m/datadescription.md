# Data description for MWC_ETP_Data.xlsx

## Number of sheets: 
5

## Sheet name: ETP Influent Flow

### Column Headings
'Date', 'Influent Flow (ML/d)'

### First 10 Rows
|    | Date                |   Influent Flow (ML/d) |
|---:|:--------------------|-----------------------:|
|  0 | 2014-01-01 00:00:00 |                    289 |
|  1 | 2014-01-02 00:00:00 |                    332 |
|  2 | 2014-01-03 00:00:00 |                    302 |
|  3 | 2014-01-04 00:00:00 |                    295 |
|  4 | 2014-01-05 00:00:00 |                    317 |
|  5 | 2014-01-06 00:00:00 |                    306 |
|  6 | 2014-01-07 00:00:00 |                    351 |
|  7 | 2014-01-08 00:00:00 |                   -209 |
|  8 | 2014-01-09 00:00:00 |                    310 |
|  9 | 2014-01-10 00:00:00 |                     98 |

### Bottom 10 Rows
|      | Date                |   Influent Flow (ML/d) |
|-----:|:--------------------|-----------------------:|
| 1816 | 2018-12-22 00:00:00 |                    404 |
| 1817 | 2018-12-23 00:00:00 |                    355 |
| 1818 | 2018-12-24 00:00:00 |                    288 |
| 1819 | 2018-12-25 00:00:00 |                    306 |
| 1820 | 2018-12-26 00:00:00 |                    292 |
| 1821 | 2018-12-27 00:00:00 |                    317 |
| 1822 | 2018-12-28 00:00:00 |                    326 |
| 1823 | 2018-12-29 00:00:00 |                    296 |
| 1824 | 2018-12-30 00:00:00 |                    287 |
| 1825 | 2018-12-31 00:00:00 |                    264 |

### Describe (include='all')
|       | Date                |   Influent Flow (ML/d) |
|:------|:--------------------|-----------------------:|
| count | 1826                |              1826      |
| mean  | 2016-07-01 12:00:00 |               347.117  |
| min   | 2014-01-01 00:00:00 |              -209      |
| 25%   | 2015-04-02 06:00:00 |               315      |
| 50%   | 2016-07-01 12:00:00 |               334      |
| 75%   | 2017-09-30 18:00:00 |               360      |
| max   | 2018-12-31 00:00:00 |              1330      |
| std   | nan                 |                68.4975 |

### Info (verbose=True)
```
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 1826 entries, 0 to 1825
Data columns (total 2 columns):
 #   Column                Non-Null Count  Dtype         
---  ------                --------------  -----         
 0   Date                  1826 non-null   datetime64[ns]
 1   Influent Flow (ML/d)  1826 non-null   int64         
dtypes: datetime64[ns](1), int64(1)
memory usage: 28.7 KB

```

### Shape
(1826, 2)

## Sheet name: ETP Influent Quality

### Column Headings
'Date', 'Ammonia (mg/L)', 'BOD (mg/L)', 'COD (mg/L)', 'Nitrate plus Nitrite (mg/L)', 'Total Nitrogen (mg/L)'

### First 10 Rows
|    | Date                |   Ammonia (mg/L) |   BOD (mg/L) |   COD (mg/L) |   Nitrate plus Nitrite (mg/L) |   Total Nitrogen (mg/L) |
|---:|:--------------------|-----------------:|-------------:|-------------:|------------------------------:|------------------------:|
|  0 | 2014-01-01 00:00:00 |               27 |          nan |          730 |                        nan    |                     nan |
|  1 | 2014-01-02 00:00:00 |               25 |          nan |          740 |                        nan    |                     nan |
|  2 | 2014-01-05 00:00:00 |               42 |          nan |          836 |                        nan    |                     nan |
|  3 | 2014-01-06 00:00:00 |               36 |          430 |          850 |                          0.01 |                      63 |
|  4 | 2014-01-07 00:00:00 |               46 |          nan |         1016 |                        nan    |                     nan |
|  5 | 2014-01-08 00:00:00 |               40 |          nan |          820 |                        nan    |                     nan |
|  6 | 2014-01-09 00:00:00 |               51 |          nan |         1110 |                        nan    |                     nan |
|  7 | 2014-01-12 00:00:00 |               41 |          nan |          730 |                        nan    |                     nan |
|  8 | 2014-01-13 00:00:00 |               26 |          nan |          710 |                        nan    |                     nan |
|  9 | 2014-01-14 00:00:00 |               42 |          530 |          830 |                        nan    |                     nan |

### Bottom 10 Rows
|      | Date                |   Ammonia (mg/L) |   BOD (mg/L) |   COD (mg/L) |   Nitrate plus Nitrite (mg/L) |   Total Nitrogen (mg/L) |
|-----:|:--------------------|-----------------:|-------------:|-------------:|------------------------------:|------------------------:|
| 1252 | 2018-12-09 00:00:00 |               56 |          330 |          920 |                        nan    |                     nan |
| 1253 | 2018-12-10 00:00:00 |               47 |          390 |          930 |                          0.01 |                      69 |
| 1254 | 2018-12-11 00:00:00 |               40 |          350 |         1100 |                        nan    |                     nan |
| 1255 | 2018-12-12 00:00:00 |               34 |          320 |          990 |                        nan    |                     nan |
| 1256 | 2018-12-13 00:00:00 |               23 |          300 |          740 |                        nan    |                     nan |
| 1257 | 2018-12-16 00:00:00 |               32 |          230 |          640 |                        nan    |                     nan |
| 1258 | 2018-12-17 00:00:00 |               30 |          290 |          750 |                        nan    |                     nan |
| 1259 | 2018-12-18 00:00:00 |               35 |          330 |          970 |                        nan    |                     nan |
| 1260 | 2018-12-19 00:00:00 |               32 |          330 |          890 |                        nan    |                     nan |
| 1261 | 2018-12-20 00:00:00 |               32 |          360 |          850 |                        nan    |                     nan |

### Describe (include='all')
|       | Date                          |   Ammonia (mg/L) |   BOD (mg/L) |   COD (mg/L) |   Nitrate plus Nitrite (mg/L) |   Total Nitrogen (mg/L) |
|:------|:------------------------------|-----------------:|-------------:|-------------:|------------------------------:|------------------------:|
| count | 1262                          |       1261       |     846      |     1225     |                   125         |               125       |
| mean  | 2016-06-27 21:04:16.735340544 |         39.3013  |     376.797  |      845.875 |                     0.01872   |                62.024   |
| min   | 2014-01-01 00:00:00           |         13       |     140      |      360     |                     0.01      |                40       |
| 25%   | 2015-04-08 06:00:00           |         34       |     320      |      750     |                     0.01      |                56       |
| 50%   | 2016-06-29 12:00:00           |         39       |     350      |      840     |                     0.01      |                62       |
| 75%   | 2017-09-20 18:00:00           |         45       |     410      |      930     |                     0.01      |                68       |
| max   | 2018-12-20 00:00:00           |         93       |     810      |     1700     |                     0.31      |                92       |
| std   | nan                           |          7.82312 |      90.0179 |      149.984 |                     0.0347335 |                 8.46288 |

### Info (verbose=True)
```
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 1262 entries, 0 to 1261
Data columns (total 6 columns):
 #   Column                       Non-Null Count  Dtype         
---  ------                       --------------  -----         
 0   Date                         1262 non-null   datetime64[ns]
 1   Ammonia (mg/L)               1261 non-null   float64       
 2   BOD (mg/L)                   846 non-null    float64       
 3   COD (mg/L)                   1225 non-null   float64       
 4   Nitrate plus Nitrite (mg/L)  125 non-null    float64       
 5   Total Nitrogen (mg/L)        125 non-null    float64       
dtypes: datetime64[ns](1), float64(5)
memory usage: 59.3 KB

```

### Shape
(1262, 6)

## Sheet name: ETP Effluent Quality

### Column Headings
'Date', 'COD (mg/L)', 'BOD (mg/L)', 'Ammonia (mg/L)', 'Total Kjeldahl Nitrogen (mg/L)', 'Total Nitrogen (mg/L)', 'Nitrate plus Nitrite (mg/L)'

### First 10 Rows
|    | Date                |   COD (mg/L) |   BOD (mg/L) |   Ammonia (mg/L) |   Total Kjeldahl Nitrogen (mg/L) |   Total Nitrogen (mg/L) |   Nitrate plus Nitrite (mg/L) |
|---:|:--------------------|-------------:|-------------:|-----------------:|---------------------------------:|------------------------:|------------------------------:|
|  0 | 2014-01-01 00:00:00 |           30 |          nan |            0.1   |                             1.5  |                     nan |                           nan |
|  1 | 2014-01-02 00:00:00 |           31 |          nan |            0.06  |                             2    |                     nan |                           nan |
|  2 | 2014-01-05 00:00:00 |           33 |          nan |            0.05  |                             4    |                     nan |                           nan |
|  3 | 2014-01-06 00:00:00 |           25 |            2 |            0.065 |                             3.15 |                      20 |                            19 |
|  4 | 2014-01-07 00:00:00 |           30 |          nan |            0.05  |                             1.8  |                     nan |                           nan |
|  5 | 2014-01-08 00:00:00 |           35 |          nan |            0.05  |                             2    |                     nan |                           nan |
|  6 | 2014-01-09 00:00:00 |           27 |          nan |            0.05  |                             1.9  |                     nan |                           nan |
|  7 | 2014-01-12 00:00:00 |           26 |          nan |            0.05  |                             3.2  |                     nan |                           nan |
|  8 | 2014-01-13 00:00:00 |           30 |          nan |            0.05  |                             1.5  |                     nan |                           nan |
|  9 | 2014-01-14 00:00:00 |           35 |            2 |            0.6   |                             4.3  |                     nan |                           nan |

### Bottom 10 Rows
|      | Date                |   COD (mg/L) |   BOD (mg/L) |   Ammonia (mg/L) |   Total Kjeldahl Nitrogen (mg/L) |   Total Nitrogen (mg/L) |   Nitrate plus Nitrite (mg/L) |
|-----:|:--------------------|-------------:|-------------:|-----------------:|---------------------------------:|------------------------:|------------------------------:|
| 1275 | 2018-12-14 00:00:00 |           28 |          nan |         0.1      |                              2.5 |                     nan |                           nan |
| 1276 | 2018-12-17 00:00:00 |           20 |          nan |         0.1      |                              4   |                     nan |                           nan |
| 1277 | 2018-12-18 00:00:00 |           15 |            2 |         0.193333 |                              3.2 |                     nan |                           nan |
| 1278 | 2018-12-19 00:00:00 |           24 |          nan |         0.1      |                              4.3 |                     nan |                           nan |
| 1279 | 2018-12-20 00:00:00 |           20 |          nan |         0.01     |                              3.7 |                     nan |                           nan |
| 1280 | 2018-12-21 00:00:00 |           15 |          nan |         0.13     |                              3.5 |                     nan |                           nan |
| 1281 | 2018-12-24 00:00:00 |            4 |          nan |       nan        |                            nan   |                     nan |                           nan |
| 1282 | 2018-12-27 00:00:00 |            2 |            2 |         0.1      |                              1.3 |                      18 |                            17 |
| 1283 | 2018-12-28 00:00:00 |            4 |          nan |       nan        |                            nan   |                     nan |                           nan |
| 1284 | 2018-12-31 00:00:00 |            6 |          nan |       nan        |                            nan   |                     nan |                           nan |

### Describe (include='all')
|       | Date                          |   COD (mg/L) |   BOD (mg/L) |   Ammonia (mg/L) |   Total Kjeldahl Nitrogen (mg/L) |   Total Nitrogen (mg/L) |   Nitrate plus Nitrite (mg/L) |
|:------|:------------------------------|-------------:|-------------:|-----------------:|---------------------------------:|------------------------:|------------------------------:|
| count | 1285                          |   1265       |   258        |      1279        |                       1267       |               129       |                     129       |
| mean  | 2016-07-08 16:56:24.280155648 |     24.683   |     2.41085  |         0.345913 |                          2.23982 |                15.4109  |                      13.8682  |
| min   | 2014-01-01 00:00:00           |      2       |     2        |         0        |                          0.5     |                 7       |                       6       |
| 25%   | 2015-04-14 00:00:00           |     20       |     2        |         0.07     |                          1.6     |                13       |                      12       |
| 50%   | 2016-07-14 00:00:00           |     25       |     2        |         0.085    |                          2       |                15       |                      14       |
| 75%   | 2017-10-10 00:00:00           |     30       |     3        |         0.1      |                          2.6     |                17       |                      16       |
| max   | 2018-12-31 00:00:00           |     80       |     6        |       316        |                         24       |                25       |                      23       |
| std   | nan                           |      6.99818 |     0.701443 |         8.8343   |                          1.11226 |                 3.05576 |                       3.08823 |

### Info (verbose=True)
```
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 1285 entries, 0 to 1284
Data columns (total 7 columns):
 #   Column                          Non-Null Count  Dtype         
---  ------                          --------------  -----         
 0   Date                            1285 non-null   datetime64[ns]
 1   COD (mg/L)                      1265 non-null   float64       
 2   BOD (mg/L)                      258 non-null    float64       
 3   Ammonia (mg/L)                  1279 non-null   float64       
 4   Total Kjeldahl Nitrogen (mg/L)  1267 non-null   float64       
 5   Total Nitrogen (mg/L)           129 non-null    float64       
 6   Nitrate plus Nitrite (mg/L)     129 non-null    float64       
dtypes: datetime64[ns](1), float64(6)
memory usage: 70.4 KB

```

### Shape
(1285, 7)

## Sheet name: ETP Electricity Usage

### Column Headings
'Date', 'Main  feeder 1 (kwh/d)', 'Main  feeder 2 (kwh/d)', 'Digester Gas Electricity Generation (kWh/d)'

### First 10 Rows
|    | Date                |   Main  feeder 1 (kwh/d) |   Main  feeder 2 (kwh/d) |   Digester Gas Electricity Generation (kWh/d) |
|---:|:--------------------|-------------------------:|-------------------------:|----------------------------------------------:|
|  0 | 2014-01-01 00:00:00 |                   105695 |                    70161 |                                        116714 |
|  1 | 2014-01-02 00:00:00 |                   157065 |                    77017 |                                         89688 |
|  2 | 2014-01-03 00:00:00 |                   109304 |                   111382 |                                        144690 |
|  3 | 2014-01-04 00:00:00 |                   150965 |                   138223 |                                        116914 |
|  4 | 2014-01-05 00:00:00 |                   152308 |                   119156 |                                        133125 |
|  5 | 2014-01-06 00:00:00 |                   129886 |                   114277 |                                        125768 |
|  6 | 2014-01-07 00:00:00 |                   170110 |                   121885 |                                        115813 |
|  7 | 2014-01-08 00:00:00 |                   172278 |                   131990 |                                        116810 |
|  8 | 2014-01-09 00:00:00 |                   168112 |                   128100 |                                        116317 |
|  9 | 2014-01-10 00:00:00 |                   184897 |                   129837 |                                         89372 |

### Bottom 10 Rows
|      | Date                |   Main  feeder 1 (kwh/d) |   Main  feeder 2 (kwh/d) |   Digester Gas Electricity Generation (kWh/d) |
|-----:|:--------------------|-------------------------:|-------------------------:|----------------------------------------------:|
| 1816 | 2018-12-22 00:00:00 |                   186917 |                   123129 |                                        124299 |
| 1817 | 2018-12-23 00:00:00 |                   195395 |                   115220 |                                        123632 |
| 1818 | 2018-12-24 00:00:00 |                   176918 |                   110713 |                                        130367 |
| 1819 | 2018-12-25 00:00:00 |                   158119 |                   100303 |                                        127991 |
| 1820 | 2018-12-26 00:00:00 |                   129703 |                    76627 |                                        120498 |
| 1821 | 2018-12-27 00:00:00 |                   163749 |                   104766 |                                        126717 |
| 1822 | 2018-12-28 00:00:00 |                   162279 |                   121680 |                                        130890 |
| 1823 | 2018-12-29 00:00:00 |                   126907 |                   106254 |                                        130739 |
| 1824 | 2018-12-30 00:00:00 |                    45782 |                   123561 |                                        129282 |
| 1825 | 2018-12-31 00:00:00 |                        0 |                   195799 |                                        135479 |

### Describe (include='all')
|       | Date                |   Main  feeder 1 (kwh/d) |   Main  feeder 2 (kwh/d) |   Digester Gas Electricity Generation (kWh/d) |
|:------|:--------------------|-------------------------:|-------------------------:|----------------------------------------------:|
| count | 1826                |                   1826   |                   1826   |                                        1826   |
| mean  | 2016-07-01 12:00:00 |                 155927   |                 118021   |                                      116839   |
| min   | 2014-01-01 00:00:00 |                      0   |                      0   |                                       19752   |
| 25%   | 2015-04-02 06:00:00 |                 136855   |                 104242   |                                      106806   |
| 50%   | 2016-07-01 12:00:00 |                 155902   |                 119820   |                                      116548   |
| 75%   | 2017-09-30 18:00:00 |                 175784   |                 136071   |                                      130739   |
| max   | 2018-12-31 00:00:00 |                 263371   |                 195799   |                                      171445   |
| std   | nan                 |                  30227.4 |                  26881.1 |                                       20402.2 |

### Info (verbose=True)
```
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 1826 entries, 0 to 1825
Data columns (total 4 columns):
 #   Column                                       Non-Null Count  Dtype         
---  ------                                       --------------  -----         
 0   Date                                         1826 non-null   datetime64[ns]
 1   Main  feeder 1 (kwh/d)                       1826 non-null   int64         
 2   Main  feeder 2 (kwh/d)                       1826 non-null   int64         
 3   Digester Gas Electricity Generation (kWh/d)  1826 non-null   int64         
dtypes: datetime64[ns](1), int64(3)
memory usage: 57.2 KB

```

### Shape
(1826, 4)

## Sheet name: Discharge Limits

### Column Headings
'Parameter', 'Discharge Limit', 'Units'

### First 10 Rows
|    | Parameter                 |   Discharge Limit | Units   |
|---:|:--------------------------|------------------:|:--------|
|  0 | Flow Rate - annual mean   |             540   | ML/d    |
|  1 | Ammonia - 90th percentile |               2   | mg/L    |
|  2 | Ammonia - annual mean     |               0.5 | mg/L    |
|  3 | BOD - 90th percentile     |              10   | mg/L    |

### Bottom 10 Rows
|    | Parameter                 |   Discharge Limit | Units   |
|---:|:--------------------------|------------------:|:--------|
|  0 | Flow Rate - annual mean   |             540   | ML/d    |
|  1 | Ammonia - 90th percentile |               2   | mg/L    |
|  2 | Ammonia - annual mean     |               0.5 | mg/L    |
|  3 | BOD - 90th percentile     |              10   | mg/L    |

### Describe (include='all')
|        | Parameter               |   Discharge Limit | Units   |
|:-------|:------------------------|------------------:|:--------|
| count  | 4                       |             4     | 4       |
| unique | 4                       |           nan     | 2       |
| top    | Flow Rate - annual mean |           nan     | mg/L    |
| freq   | 1                       |           nan     | 3       |
| mean   | nan                     |           138.125 | nan     |
| std    | nan                     |           267.949 | nan     |
| min    | nan                     |             0.5   | nan     |
| 25%    | nan                     |             1.625 | nan     |
| 50%    | nan                     |             6     | nan     |
| 75%    | nan                     |           142.5   | nan     |
| max    | nan                     |           540     | nan     |

### Info (verbose=True)
```
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 4 entries, 0 to 3
Data columns (total 3 columns):
 #   Column           Non-Null Count  Dtype  
---  ------           --------------  -----  
 0   Parameter        4 non-null      object 
 1   Discharge Limit  4 non-null      float64
 2   Units            4 non-null      object 
dtypes: float64(1), object(2)
memory usage: 228.0+ bytes

```

### Shape
(4, 3)

### Unique values in text columns (up to 100 values)
#### Column: Parameter
'Flow Rate - annual mean', 'Ammonia - 90th percentile', 'Ammonia - annual mean', 'BOD - 90th percentile'

#### Column: Units
'ML/d', 'mg/L'

