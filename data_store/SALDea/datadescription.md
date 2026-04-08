# Data description for WWTP_microbial_loads_and_removal.xlsx

## Number of sheets: 
3

## Sheet name: Data

### Column Headings
'Event', 'Plant', 'SamplingMonth', 'SampleType', 'IndicatorType', 'IndicatorName', 'Count'

### First 10 Rows
|    | Event    | Plant   | SamplingMonth   | SampleType    | IndicatorType   | IndicatorName          |   Count |
|---:|:---------|:--------|:----------------|:--------------|:----------------|:-----------------------|--------:|
|  0 | Event 01 | NoWe    | Mar 2015        | Influent grab | Coliphage       | SomaColiphage_PFUperML |    3050 |
|  1 | Event 01 | NoWe    | Mar 2015        | Effluent grab | Coliphage       | SomaColiphage_PFUperML |       0 |
|  2 | Event 01 | NoWe    | Mar 2015        | UF Effluent   | Coliphage       | SomaColiphage_PFUperML |       0 |
|  3 | Event 01 | BiSp    | Mar 2015        | Influent grab | Coliphage       | SomaColiphage_PFUperML |    1300 |
|  4 | Event 01 | BiSp    | Mar 2015        | Effluent grab | Coliphage       | SomaColiphage_PFUperML |       2 |
|  5 | Event 01 | BiSp    | Mar 2015        | UF Effluent   | Coliphage       | SomaColiphage_PFUperML |      66 |
|  6 | Event 02 | BrWo    | Apr 2015        | Influent grab | Coliphage       | SomaColiphage_PFUperML |   15150 |
|  7 | Event 02 | BrWo    | Apr 2015        | Effluent grab | Coliphage       | SomaColiphage_PFUperML |       0 |
|  8 | Event 02 | BrWo    | Apr 2015        | UF Effluent   | Coliphage       | SomaColiphage_PFUperML |       0 |
|  9 | Event 02 | NoWe    | Apr 2015        | Influent grab | Coliphage       | SomaColiphage_PFUperML |     615 |

### Bottom 10 Rows
|      | Event    | Plant   | SamplingMonth   | SampleType    | IndicatorType   | IndicatorName      | Count   |
|-----:|:---------|:--------|:----------------|:--------------|:----------------|:-------------------|:--------|
| 1214 | Event 15 | BiSP    | May 2016        | Influent grab | Virus           | Adenovirus (MPN/L) | 17      |
| 1215 | Event 15 | BiSP    | May 2016        | Effluent grab | Virus           | Adenovirus (MPN/L) | 0.09    |
| 1216 | Event 15 | BrWo    | May 2016        | Influent grab | Virus           | Adenovirus (MPN/L) | 17      |
| 1217 | Event 15 | BrWo    | May 2016        | Effluent grab | Virus           | Adenovirus (MPN/L) | BDL     |
| 1218 | Event 16 | NoWe    | Jun 2016        | Influent grab | Virus           | Adenovirus (MPN/L) | 27975   |
| 1219 | Event 16 | NoWe    | Jun 2016        | Effluent grab | Virus           | Adenovirus (MPN/L) | 511     |
| 1220 | Event 16 | BiSP    | Jun 2016        | Influent grab | Virus           | Adenovirus (MPN/L) | 20      |
| 1221 | Event 16 | BiSP    | Jun 2016        | Effluent grab | Virus           | Adenovirus (MPN/L) | 0.05    |
| 1222 | Event 16 | BrWo    | Jun 2016        | Influent grab | Virus           | Adenovirus (MPN/L) | 143     |
| 1223 | Event 16 | BrWo    | Jun 2016        | Effluent grab | Virus           | Adenovirus (MPN/L) | BDL     |

### Describe (include='all')
|        | Event    | Plant   | SamplingMonth   | SampleType    | IndicatorType   | IndicatorName          |   Count |
|:-------|:---------|:--------|:----------------|:--------------|:----------------|:-----------------------|--------:|
| count  | 1224     | 1224    | 1224            | 1224          | 1224            | 1224                   |     906 |
| unique | 16       | 4       | 16              | 3             | 4               | 9                      |     470 |
| top    | Event 02 | NoWe    | Apr 2015        | Influent grab | Bacteria        | SomaColiphage_PFUperML |       0 |
| freq   | 78       | 416     | 78              | 424           | 564             | 141                    |     192 |

### Info (verbose=True)
```
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 1224 entries, 0 to 1223
Data columns (total 7 columns):
 #   Column         Non-Null Count  Dtype 
---  ------         --------------  ----- 
 0   Event          1224 non-null   object
 1   Plant          1224 non-null   object
 2   SamplingMonth  1224 non-null   object
 3   SampleType     1224 non-null   object
 4   IndicatorType  1224 non-null   object
 5   IndicatorName  1224 non-null   object
 6   Count          906 non-null    object
dtypes: object(7)
memory usage: 67.1+ KB

```

### Shape
(1224, 7)

### Unique values in text columns (non-numeric only, up to 100 values)
#### Column: Event
'Event 01', 'Event 02', 'Event 03', 'Event 04', 'Event 05', 'Event 06', 'Event 07', 'Event 08', 'Event 09', 'Event 10', 'Event 11', 'Event 12', 'Event 13', 'Event 14', 'Event 15', 'Event 16'

#### Column: Plant
'NoWe', 'BiSp', 'BrWo', 'BiSP'

#### Column: SamplingMonth
'Mar 2015', 'Apr 2015', 'May 2015', 'Jun 2015', 'Jul 2015', 'Aug 2015', 'Sept 2015', 'Oct 2015', 'Nov 2015', 'Dec 2015', 'Jan 2016', 'Feb 2016', 'Mar 2016', 'Apr 2016', 'May 2016', 'Jun 2016'

#### Column: SampleType
'Influent grab', 'Effluent grab', 'UF Effluent'

#### Column: IndicatorType
'Coliphage', 'Bacteria', 'Parasite', 'Virus'

#### Column: IndicatorName
'SomaColiphage_PFUperML', 'MaleSpecColiphage_PFUperML', 'AerobicEndospore_CFUperML', 'TotalColiform_MPNper100ML', 'FecalColiform_MPNper100ML', 'E.coli_MPNper100ML', 'CryptoOocyst_perL', 'GiardiaCyst_perL', 'Adenovirus (MPN/L)'

#### Column: Count
'BDL', '<NaN>'

## Sheet name: Method detection limits

### Column Headings
'Organism', 'MDL', 'Units'

### First 10 Rows
|    | Organism                |   MDL | Units     |
|---:|:------------------------|------:|:----------|
|  0 | Adenovirus              |  0.02 | MPN/L     |
|  1 | Total coliform          |  1    | MPN/100mL |
|  2 | Fecal coliform          |  1    | MPN/100mL |
|  3 | E. coli                 |  1    | MPN/100mL |
|  4 | Aerobic endospores      |  1    | CFU/100mL |
|  5 | Soma Coliphage          |  1    | PFU/mL    |
|  6 | Male Specific Coliphage |  1    | PFU/mL    |
|  7 | Crypto Oocyst           |  0.1  | /L        |
|  8 | Giardia Cyst            |  0.1  | /L        |

### Bottom 10 Rows
|    | Organism                |   MDL | Units     |
|---:|:------------------------|------:|:----------|
|  0 | Adenovirus              |  0.02 | MPN/L     |
|  1 | Total coliform          |  1    | MPN/100mL |
|  2 | Fecal coliform          |  1    | MPN/100mL |
|  3 | E. coli                 |  1    | MPN/100mL |
|  4 | Aerobic endospores      |  1    | CFU/100mL |
|  5 | Soma Coliphage          |  1    | PFU/mL    |
|  6 | Male Specific Coliphage |  1    | PFU/mL    |
|  7 | Crypto Oocyst           |  0.1  | /L        |
|  8 | Giardia Cyst            |  0.1  | /L        |

### Describe (include='all')
|        | Organism   |        MDL | Units     |
|:-------|:-----------|-----------:|:----------|
| count  | 9          |   9        | 9         |
| unique | 9          | nan        | 5         |
| top    | Adenovirus | nan        | MPN/100mL |
| freq   | 1          | nan        | 3         |
| mean   | nan        |   0.691111 | nan       |
| std    | nan        |   0.463909 | nan       |
| min    | nan        |   0.02     | nan       |
| 25%    | nan        |   0.1      | nan       |
| 50%    | nan        |   1        | nan       |
| 75%    | nan        |   1        | nan       |
| max    | nan        |   1        | nan       |

### Info (verbose=True)
```
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 9 entries, 0 to 8
Data columns (total 3 columns):
 #   Column    Non-Null Count  Dtype  
---  ------    --------------  -----  
 0   Organism  9 non-null      object 
 1   MDL       9 non-null      float64
 2   Units     9 non-null      object 
dtypes: float64(1), object(2)
memory usage: 348.0+ bytes

```

### Shape
(9, 3)

### Unique values in text columns (non-numeric only, up to 100 values)
#### Column: Organism
'Adenovirus', 'Total coliform', 'Fecal coliform', 'E. coli', 'Aerobic endospores', 'Soma Coliphage', 'Male Specific Coliphage', 'Crypto Oocyst', 'Giardia Cyst'

#### Column: Units
'MPN/L', 'MPN/100mL', 'CFU/100mL', 'PFU/mL', '/L'

## Sheet name: Plants

### Column Headings
'WWTP Plant', 'Short Name', 'Full Name', 'Primary and Seconary Treatment Train', 'Disinfection'

### First 10 Rows
|    |   WWTP Plant | Short Name   | Full Name         | Primary and Seconary Treatment Train                                                        | Disinfection   |
|---:|-------------:|:-------------|:------------------|:--------------------------------------------------------------------------------------------|:---------------|
|  0 |            1 | NoWe         | Northwest El Paso | Primary clarification, activated sludge, secondary clarification,  disc filter microscreens | UV             |
|  1 |            2 | BiSp         | Big Spring        | Primary clarification, activated sludge/trickling filter, secondary clarification           | Chlorine       |
|  2 |            3 | BrWo         | Brownwood         | Primary clarification, activated sludge, secondary clarification                            | Chlorine       |

### Bottom 10 Rows
|    |   WWTP Plant | Short Name   | Full Name         | Primary and Seconary Treatment Train                                                        | Disinfection   |
|---:|-------------:|:-------------|:------------------|:--------------------------------------------------------------------------------------------|:---------------|
|  0 |            1 | NoWe         | Northwest El Paso | Primary clarification, activated sludge, secondary clarification,  disc filter microscreens | UV             |
|  1 |            2 | BiSp         | Big Spring        | Primary clarification, activated sludge/trickling filter, secondary clarification           | Chlorine       |
|  2 |            3 | BrWo         | Brownwood         | Primary clarification, activated sludge, secondary clarification                            | Chlorine       |

### Describe (include='all')
|        |   WWTP Plant | Short Name   | Full Name         | Primary and Seconary Treatment Train                                                        | Disinfection   |
|:-------|-------------:|:-------------|:------------------|:--------------------------------------------------------------------------------------------|:---------------|
| count  |          3   | 3            | 3                 | 3                                                                                           | 3              |
| unique |        nan   | 3            | 3                 | 3                                                                                           | 2              |
| top    |        nan   | NoWe         | Northwest El Paso | Primary clarification, activated sludge, secondary clarification,  disc filter microscreens | Chlorine       |
| freq   |        nan   | 1            | 1                 | 1                                                                                           | 2              |
| mean   |          2   | nan          | nan               | nan                                                                                         | nan            |
| std    |          1   | nan          | nan               | nan                                                                                         | nan            |
| min    |          1   | nan          | nan               | nan                                                                                         | nan            |
| 25%    |          1.5 | nan          | nan               | nan                                                                                         | nan            |
| 50%    |          2   | nan          | nan               | nan                                                                                         | nan            |
| 75%    |          2.5 | nan          | nan               | nan                                                                                         | nan            |
| max    |          3   | nan          | nan               | nan                                                                                         | nan            |

### Info (verbose=True)
```
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 3 entries, 0 to 2
Data columns (total 5 columns):
 #   Column                                Non-Null Count  Dtype 
---  ------                                --------------  ----- 
 0   WWTP Plant                            3 non-null      int64 
 1   Short Name                            3 non-null      object
 2   Full Name                             3 non-null      object
 3   Primary and Seconary Treatment Train  3 non-null      object
 4   Disinfection                          3 non-null      object
dtypes: int64(1), object(4)
memory usage: 252.0+ bytes

```

### Shape
(3, 5)

### Unique values in text columns (non-numeric only, up to 100 values)
#### Column: Short Name
'NoWe', 'BiSp', 'BrWo'

#### Column: Full Name
'Northwest El Paso', 'Big Spring', 'Brownwood'

#### Column: Primary and Seconary Treatment Train
'Primary clarification, activated sludge, secondary clarification,  disc filter microscreens', 'Primary clarification, activated sludge/trickling filter, secondary clarification', 'Primary clarification, activated sludge, secondary clarification'

#### Column: Disinfection
'UV', 'Chlorine'

