# Cadastral-Digital-Framework

यह रिपॉज़िटरी **Maharashtra Cadastral Digital Framework (SOP-MCDF-V1.0)** के अनुसार कैडस्ट्रल डेटा की जाँच (Quality Control) और एरिया रिकंसीडेशन रिपोर्ट तैयार करने से जुड़े टूल्स और स्क्रिप्ट्स के लिए है[cite: 1, 2, 3]।

## मुख्य उद्देश्य (Overview)
- गाँव-वार (Village-wise) निर्देशिका संरचना और आवश्यक फ़ाइलों की जाँच करना[cite: 2]।
- कोऑर्डिनेट सिस्टम (WGS 84 / UTM Zone 43N या 44N) और एट्रीब्यूट स्कीमा का सत्यापन करना[cite: 2, 3]।
- सरकारी 7/12 क्षेत्रफल (`ROR_AREA`) और GIS मैप क्षेत्रफल (`GIS_AREA`) की तुलना करके अंतर (`AREA_DIFF`, `DIFF_PERC`) निकालना[cite: 3]।
- पार्सल स्तर पर `MATCHED`, `ACCEPTABLE`, और `DEVIANT` स्थिति के साथ एक्सेल रिपोर्ट (`Village_Summary` और `Parcel_Level_Comparison`) तैयार करना[cite: 3]।

## आवश्यक लाइब्रेरीज़ (Prerequisites)
- Python (3.9 या उससे ऊपर)
- `geopandas`, `pandas`, `shapely`, `openpyxl`

## उपयोग के निर्देश (How to Use)
1. संबंधित गाँव की शेपफाइल (`.shp`) और 7/12 मिलान डेटा तैयार रखें[cite: 2, 3]।
2. Jupyter Notebook या Python स्क्रिप्ट में इनपुट पाथ और प्रशासनिक विवरण (District, Taluka, Village, LGD Code आदि) सेट करें[cite: 2, 3]।
3. स्क्रिप्ट चलाकर वैलिडेशन और एक्सेल रिपोर्ट जनरेट करें[cite: 3]।

## नोट (Note)
- किसी भी वास्तविक 7/12 डेटा, गोपनीय रिकॉर्ड या भारी शेपफ़ाइलों को सार्वजनिक रिपॉज़िटरी में अपलोड न करें।
