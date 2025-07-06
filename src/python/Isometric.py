# Offset to avoid having 0 lenght
DefaultOffset = 32

def InvertGridDirection1D(coord, offset, Points=None):
    # Invert the Y direction based on the maximum Y value (offset)
    if Points is not None:
        for j in range(len(Points)):
            Points[j]["Y"] = -Points[j]["Y"] + offset
        return Points
    else:
        return -coord + offset

def ShiftGrid1D(coord, offset, Points=None):
    # Shift the grid by a offset
    if Points is not None:
        for j in range(len(Points)):
            Points[j]["X"] = Points[j]["X"] - offset["X"]
            Points[j]["Y"] = Points[j]["Y"] - offset["Y"]
        return Points
    else:
        return coord - offset

def InvertGridDirection2D(coord, offset, Points=None):
    # Invert the Y direction based on the maximum Y value (offset)
    if Points is not None:
        for j in range(len(Points)):
            for i in range(len(Points[j])):
                Points[j][i]["Y"] = -Points[j][i]["Y"] + offset
        return Points
    else:
        return -coord + offset

def ShiftGrid2D(coord, offset, Points=None):
    # Shift the grid by a offset
    if Points is not None:
        for j in range(len(Points)):
            for i in range(len(Points[j])):
                Points[j][i]["X"] = Points[j][i]["X"] - offset["X"]
                Points[j][i]["Y"] = Points[j][i]["Y"] - offset["Y"]
        return Points
    else:
        return coord - offset

Points = [
    # Execute o comando LIST nas polylines do isométrico para extrair os valores.
    # O comando LIST entrega os valores ordenados da direita para esqueda em relação
    # à coordenada X.
    # Segmento 1
    [
        {"X": 2146.2196, "Y": 1599.3516, "Lenght": 150, "Diameter": "1 inch"},
        {"X": 2144.4726, "Y": 1592.2910, "Lenght": 790, "Diameter": "1 inch"},
        {"X": 2144.4726, "Y": 1569.5606, "Lenght": 3460, "Diameter": "1 inch"},
        {"X": 2137.6525, "Y": 1541.9976, "Lenght": 405, "Diameter": "1 inch"},
        {"X": 2162.0935, "Y": 1541.0569, "Lenght": 1620, "Diameter": "1 inch"},
        {"X": 2158.2469, "Y": 1525.5112, "Lenght": 4790, "Diameter": "1 inch"},
        {"X": 2158.2355, "Y": 1567.4770, "Lenght": 2140, "Diameter": "1 inch"},
        {"X": 2154.3890, "Y": 1551.9314, "Lenght": 440, "Diameter": "1 inch"},
        {"X": 2154.3890, "Y": 1561.9314, "Lenght": 11550, "Diameter": "1 inch"},
        {"X": 2203.1029, "Y": 1560.0564, "Lenght": 455, "Diameter": "1 inch"},
        {"X": 2200.9378, "Y": 1551.3064, "Lenght": 4160, "Diameter": "1 inch"},
        {"X": 2200.9378, "Y": 1511.3064, "Lenght": 1200, "Diameter": "1 inch"},
        {"X": 2213.9282, "Y": 1503.8064, "Lenght": 4765, "Diameter": "1 inch"},
        {"X": 2179.2872, "Y": 1483.8064, "Lenght": 8255, "Diameter": "1 inch"},
        {"X": 2179.2872, "Y": 1458.8064, "Lenght": 530, "Diameter": "1 inch"},
        {"X": 2194.0096, "Y": 1450.3064, "Lenght": 1090, "Diameter": "1 inch"},
        {"X": 2211.3081, "Y": 1460.3191, "Lenght": 270, "Diameter": "1 inch"},
        {"X": 2211.3081, "Y": 1469.3191, "Lenght": 11250, "Diameter": "1 inch"},
        {"X": 2297.4533, "Y": 1419.5831, "Lenght": 160, "Diameter": "1 inch"},
        {"X": 2297.4533, "Y": 1428.5831, "Lenght": 1230, "Diameter": "1 inch"},
        {"X": 2280.1328, "Y": 1418.5831, "Lenght": 160, "Diameter": "1 inch"},
        {"X": 2280.1328, "Y": 1409.5831, "Lenght": 1935, "Diameter": "1 inch"},
        {"X": 2280.1328, "Y": 1393.5831, "Lenght": 4610, "Diameter": "1 inch"},
        {"X": 2343.3130, "Y": 1357.1060, "Lenght": 1115, "Diameter": "1 inch"},
        {"X": 2446.3700, "Y": 1297.6060, "Lenght": 5250, "Diameter": "1 inch"},
        {"X": 2394.4085, "Y": 1267.6060, "Lenght": 1950, "Diameter": "1 inch"},
        {"X": 2394.4085, "Y": 1286.6060, "Lenght": 2550, "Diameter": "1 inch"},
        {"X": 2430.7815, "Y": 1265.6060, "Lenght": 1060, "Diameter": "1 inch"},
        {"X": 2430.7815, "Y": 1274.6060, "Lenght": 360, "Diameter": "1 inch"},
        {"X": 2446.3700, "Y": 1283.6060, "Lenght": 12630, "Diameter": "1 inch"},
        {"X": 2560.6853, "Y": 1217.6060, "Lenght": 335, "Diameter": "1 inch"},
        {"X": 2576.2738, "Y": 1226.6060, "Lenght": 120, "Diameter": "1 inch"},
        {"X": 2586.6661, "Y": 1220.6060, "Lenght": 3235, "Diameter": "1 inch"},
        {"X": 2586.6661, "Y": 1191.6060, "Lenght": 6495, "Diameter": "1 inch"},
        {"X": 2629.9674, "Y": 1166.6060, "Lenght": 6805, "Diameter": "1 inch"},
        {"X": 2673.2686, "Y": 1141.6060, "Lenght": 14550, "Diameter": "1 inch"},
        {"X": 2716.5699, "Y": 1116.6060, "Lenght": 190, "Diameter": "1 inch"},
        {"X": 2716.5699, "Y": 1130.6060, "Lenght": 820, "Diameter": "1 inch"},
        {"X": 2728.6943, "Y": 1137.6060, "Lenght": 290, "Diameter": "1 inch"},
        {"X": 2728.6943, "Y": 1123.6060, "Lenght": 430, "Diameter": "1 inch"},
        {"X": 2745.1487, "Y": 1133.1060, "Lenght": 1910, "Diameter": "1 inch"},
        {"X": 2745.1487, "Y": 1157.1060, "Lenght": 130, "Diameter": "1 inch"},
        {"X": 2736.4885, "Y": 1162.1060, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 2
    [
        {"X": 2280.1328, "Y": 1409.5831, "Lenght": 3570, "Diameter": "1 inch"},
        {"X": 2246.3578, "Y": 1390.0831, "Lenght": 1290, "Diameter": "1 inch"},
        {"X": 2246.3578, "Y": 1366.0831, "Lenght": 764, "Diameter": "1 inch"},
        {"X": 2222.9751, "Y": 1379.5831, "Lenght": 1750, "Diameter": "1 inch"},
        {"X": 2222.9751, "Y": 1406.5831, "Lenght": 1, "Diameter": "1 inch"},
        {"X": 2219.9440, "Y": 1404.8331, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Semgnto 3
    [
        {"X": 2343.3130, "Y": 1357.1060, "Lenght": 2450, "Diameter": "1 inch"},
        {"X": 2343.3130, "Y": 1471.1433, "Lenght": 3905, "Diameter": "1 inch"},
        {"X": 2406.8204, "Y": 1507.8093, "Lenght": 4530, "Diameter": "1 inch"},
        {"X": 2453.9647, "Y": 1505.9947, "Lenght": 5315, "Diameter": "1 inch"},
        {"X": 2492.9359, "Y": 1528.4947, "Lenght": 4045, "Diameter": "1 inch"},
        {"X": 2540.0803, "Y": 1526.6801, "Lenght": 430, "Diameter": "1 inch"},
        {"X": 2540.0803, "Y": 1517.6801, "Lenght": 650, "Diameter": "1 inch"},
        {"X": 2556.5347, "Y": 1527.1801, "Lenght": 160, "Diameter": "1 inch"},
        {"X": 2556.5347, "Y": 1518.1801, "Lenght": 1060, "Diameter": "1 inch"},
        {"X": 2581.6495, "Y": 1503.6801, "Lenght": 5740, "Diameter": "1 inch"},
        {"X": 2697.6969, "Y": 1436.6801, "Lenght": 870, "Diameter": "1 inch"},
        {"X": 2709.8212, "Y": 1443.6801, "Lenght": 2550, "Diameter": "1 inch"},
        {"X": 2753.1225, "Y": 1468.6801, "Lenght": 5030, "Diameter": "1 inch"},
        {"X": 2841.4571, "Y": 1417.6801, "Lenght": 2110, "Diameter": "1 inch"},
        {"X": 2841.4571, "Y": 1390.6801, "Lenght": 290, "Diameter": "1 inch"},
        {"X": 2857.9116, "Y": 1381.1801, "Lenght": 535, "Diameter": "1 inch"},
        {"X": 2887.3564, "Y": 1364.1801, "Lenght": 2035, "Diameter": "1 inch"},
        {"X": 2925.4616, "Y": 1386.1801, "Lenght": 8095, "Diameter": "1 inch"},
        {"X": 2972.2269, "Y": 1359.1801, "Lenght": 13300, "Diameter": "1 inch"},
        {"X": 3040.6429, "Y": 1319.6801, "Lenght": 100, "Diameter": "1 inch"},
        {"X": 3040.6429, "Y": 1328.6801, "Lenght": 300, "Diameter": "1 inch"},
        {"X": 3052.7673, "Y": 1335.6801, "Lenght": 1590, "Diameter": "1 inch"},
        {"X": 3052.7673, "Y": 1359.6801, "Lenght": 150, "Diameter": "1 inch"},
        {"X": 3060.9945, "Y": 1354.9301, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 4
    [
        {"X": 2581.6495, "Y": 1503.6801, "Lenght": 150, "Diameter": "1 inch"},
        {"X": 2581.6495, "Y": 1517.6801, "Lenght": 210, "Diameter": "1 inch"},
        {"X": 2593.7738, "Y": 1524.6801, "Lenght": 1640, "Diameter": "1 inch"},
        {"X": 2593.7738, "Y": 1472.6801, "Lenght": 2355, "Diameter": "1 inch"},
        {"X": 2641.4052, "Y": 1500.1801, "Lenght": 3820, "Diameter": "1 inch"},
        {"X": 2598.1040, "Y": 1525.1801, "Lenght": 1780, "Diameter": "1 inch"},
        {"X": 2621.1362, "Y": 1538.4778, "Lenght": 2164, "Diameter": "1 inch"},
        {"X": 2621.1362, "Y": 1568.4778, "Lenght": 280, "Diameter": "1 inch"},
        {"X": 2614.6410, "Y": 1572.2278, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 5
    [
        {"X": 2709.8212, "Y": 1443.6801, "Lenght": 120, "Diameter": "1 inch"},
        {"X": 2709.8212, "Y": 1455.1801, "Lenght": 300, "Diameter": "1 inch"},
        {"X": 2699.4289, "Y": 1461.1801, "Lenght": 10370, "Diameter": "1 inch"},
        {"X": 2699.4289, "Y": 1503.1801, "Lenght": 400, "Diameter": "1 inch"},
        {"X": 2690.7687, "Y": 1498.1801, "Lenght": 3940, "Diameter": "1 inch"},
        {"X": 2654.3956, "Y": 1519.1801, "Lenght": 2230, "Diameter": "1 inch"},
        {"X": 2654.3956, "Y": 1548.1801, "Lenght": 280, "Diameter": "1 inch"},
        {"X": 2662.6229, "Y": 1543.4301, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 6
    [
        {"X": 2857.9116, "Y": 1381.1801, "Lenght": 160, "Diameter": "1 inch"},
        {"X": 2857.9116, "Y": 1390.1801, "Lenght": 945, "Diameter": "1 inch"},
        {"X": 2837.1270, "Y": 1378.1801, "Lenght": 140, "Diameter": "1 inch"},
        {"X": 2837.1270, "Y": 1369.1801, "Lenght": 2735, "Diameter": "1 inch"},
        {"X": 2792.0937, "Y": 1343.1801, "Lenght": 1230, "Diameter": "1 inch"},
        {"X": 2762.6488, "Y": 1326.1801, "Lenght": 1280, "Diameter": "1 inch"},
        {"X": 2740.9982, "Y": 1338.6801, "Lenght": 1070, "Diameter": "1 inch"},
        {"X": 2740.9982, "Y": 1362.6801, "Lenght": 150, "Diameter": "1 inch"},
        {"X": 2734.0699, "Y": 1358.6801, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 7
    [
        {"X": 2999.7657, "Y": 1093.1029, "Lenght": 2460, "Diameter": "1 inch"},
        {"X": 2999.7657, "Y": 1071.1425, "Lenght": 650, "Diameter": "1 inch"},
        {"X": 3032.6747, "Y": 1052.1029, "Lenght": 1520, "Diameter": "1 inch"},
        {"X": 3032.6747, "Y": 1023.1029, "Lenght": 850, "Diameter": "1 inch"},
        {"X": 3049.1292, "Y": 1013.6029, "Lenght": 3140, "Diameter": "1 inch"},
        {"X": 3100.2247, "Y": 1043.1029, "Lenght": 2800, "Diameter": "1 inch"},
        {"X": 3138.3298, "Y": 1065.1029, "Lenght": 950, "Diameter": "1 inch"},
        {"X": 3117.5452, "Y": 1077.1425, "Lenght": 1985, "Diameter": "1 inch"},
        {"X": 3117.5452, "Y": 1106.1425, "Lenght": 1985, "Diameter": "1 inch"},
        {"X": 3117.5452, "Y": 1135.1425, "Lenght": 3460, "Diameter": "1 inch"},
        {"X": 3070.7798, "Y": 1162.1425, "Lenght": 1190, "Diameter": "1 inch"},
        {"X": 3036.6547, "Y": 1181.8447, "Lenght": 4660, "Diameter": "1 inch"},
        {"X": 2975.5170, "Y": 1217.1425, "Lenght": 4660, "Diameter": "1 inch"},
        {"X": 2932.2158, "Y": 1242.1425, "Lenght": 4650, "Diameter": "1 inch"},
        {"X": 2888.9145, "Y": 1267.1425, "Lenght": 4650, "Diameter": "1 inch"},
        {"X": 2845.6132, "Y": 1292.1425, "Lenght": 3720, "Diameter": "1 inch"},
        {"X": 2804.0440, "Y": 1316.1425, "Lenght": 2060, "Diameter": "1 inch"},
        {"X": 2804.0440, "Y": 1284.1425, "Lenght": 910, "Diameter": "1 inch"},
        {"X": 2804.0440, "Y": 1272.1425, "Lenght": 250, "Diameter": "1 inch"},
        {"X": 2795.3837, "Y": 1257.1425, "Lenght": 930, "Diameter": "1 inch"},
        {"X": 2795.3837, "Y": 1243.1425, "Lenght": 535, "Diameter": "1 inch"},
        {"X": 2825.8142, "Y": 1260.7116, "Lenght": 2135, "Diameter": "1 inch"},
        {"X": 2773.0356, "Y": 1291.1833, "Lenght": 200, "Diameter": "1 inch"},
        {"X": 2773.0356, "Y": 1300.1833, "Lenght": 3280, "Diameter": "1 inch"},
        {"X": 2819.8009, "Y": 1327.1833, "Lenght": 990, "Diameter": "1 inch"},
        {"X": 2792.0937, "Y": 1343.1801, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 8
    [
        {"X": 3041.3350, "Y": 1051.1029, "Lenght": 190, "Diameter": "1 inch"},
        {"X": 3041.3350, "Y": 1083.1029, "Lenght": 820, "Diameter": "1 inch"},
        {"X": 3029.2106, "Y": 1076.1029, "Lenght": 290, "Diameter": "1 inch"},
        {"X": 2999.7657, "Y": 1093.1029, "Lenght": 9330, "Diameter": "1 inch"},
        {"X": 2973.7850, "Y": 1108.1029, "Lenght": 515, "Diameter": "1 inch"},
        {"X": 2927.0196, "Y": 1135.1029, "Lenght": 1240, "Diameter": "1 inch"},
        {"X": 2880.2542, "Y": 1162.1029, "Lenght": 3260, "Diameter": "1 inch"},
        {"X": 2833.4889, "Y": 1189.1029, "Lenght": 2060, "Diameter": "1 inch"},
        {"X": 2786.7235, "Y": 1216.1029, "Lenght": 3370, "Diameter": "1 inch"},
        {"X": 2739.9581, "Y": 1243.1029, "Lenght": 4640, "Diameter": "1 inch"},
        {"X": 2693.1927, "Y": 1270.1029, "Lenght": 4810, "Diameter": "1 inch"},
        {"X": 2693.1927, "Y": 1246.1029, "Lenght": 4450, "Diameter": "1 inch"},
        {"X": 2693.1927, "Y": 1222.1029, "Lenght": 4620, "Diameter": "1 inch"},
        {"X": 2709.6417, "Y": 1212.6060, "Lenght": 1290, "Diameter": "1 inch"},
        {"X": 2693.1872, "Y": 1203.1060, "Lenght": 2910, "Diameter": "1 inch"},
        {"X": 2642.0917, "Y": 1173.6060, "Lenght": 1070, "Diameter": "1 inch"},
        {"X": 2642.0917, "Y": 1187.6060, "Lenght": 125, "Diameter": "1 inch"},
        {"X": 2629.9674, "Y": 1180.6060, "Lenght": 1950, "Diameter": "1 inch"},
        {"X": 2629.9674, "Y": 1166.6060, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 9
    [
        {"X": 3331.4535, "Y": 1189.7902, "Lenght": 1600+30445, "Diameter": "1 inch"},
        {"X": 3331.4535, "Y": 1084.1029, "Lenght": 17520, "Diameter": "1 inch"},
        {"X": 3156.5163, "Y":  983.1029, "Lenght": 2180, "Diameter": "1 inch"},
        {"X": 3135.7317, "Y":  995.1029, "Lenght": 3100, "Diameter": "1 inch"},
        {"X": 3135.7317, "Y": 1024.1029, "Lenght": 540, "Diameter": "1 inch"},
        {"X": 3119.2772, "Y": 1033.6029, "Lenght": 80, "Diameter": "1 inch"},
        {"X": 3110.6170, "Y": 1048.6029, "Lenght": 300, "Diameter": "1 inch"},
        {"X": 3100.2247, "Y": 1054.6029, "Lenght": 130, "Diameter": "1 inch"},
        {"X": 3100.2247, "Y": 1043.1029, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 10
    [
        {"X": 3226.2314, "Y": 1267.5402, "Lenght": 200, "Diameter": "1 inch"},
        {"X": 3232.7266, "Y": 1263.7902, "Lenght": 1250, "Diameter": "1 inch"},
        {"X": 3232.7266, "Y": 1233.7902, "Lenght": 7835, "Diameter": "1 inch"},
        {"X": 3256.1093, "Y": 1247.2902, "Lenght": 3110, "Diameter": "1 inch"},
        {"X": 3318.4631, "Y": 1211.2902, "Lenght": 520, "Diameter": "1 inch"},
        {"X": 3306.3387, "Y": 1204.2902, "Lenght": 7420, "Diameter": "1 inch"},
        {"X": 3331.4535, "Y": 1189.7902, "Lenght": 575, "Diameter": "1 inch"},
        {"X": 3399.8695, "Y": 1150.2902, "Lenght": 1250, "Diameter": "1 inch"},
        {"X": 3399.8695, "Y": 1174.2902, "Lenght": 180, "Diameter": "1 inch"},
        {"X": 3408.0967, "Y": 1169.5402, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 11
    [
        {"X": 2693.1872, "Y": 1187.1060, "Lenght": 190, "Diameter": "1 inch"},
        {"X": 2701.8475, "Y": 1182.1060, "Lenght": 820, "Diameter": "1 inch"},
        {"X": 2701.8475, "Y": 1158.1060, "Lenght": 290, "Diameter": "1 inch"},
        {"X": 2685.3930, "Y": 1148.6060, "Lenght": 430, "Diameter": "1 inch"},
        {"X": 2685.3930, "Y": 1162.6060, "Lenght": 1910, "Diameter": "1 inch"},
        {"X": 2673.2686, "Y": 1155.6060, "Lenght": 130, "Diameter": "1 inch"},
        {"X": 2673.2686, "Y": 1141.6060, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 12
    [
        {"X": 2972.2269, "Y": 1359.1801, "Lenght": 100, "Diameter": "1 inch"},
        {"X": 2972.2269, "Y": 1368.1801, "Lenght": 300, "Diameter": "1 inch"},
        {"X": 2984.3513, "Y": 1375.1801, "Lenght": 1590, "Diameter": "1 inch"},
        {"X": 2984.3513, "Y": 1399.1801, "Lenght": 150, "Diameter": "1 inch"},
        {"X": 2992.5785, "Y": 1394.4301, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 13
    [
        {"X": 2693.1927, "Y": 1246.1029, "Lenght": 160, "Diameter": "1 inch"},
        {"X": 2698.6842, "Y": 1249.2733, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 14
    [
        {"X": 2693.1872, "Y": 1203.1060, "Lenght": 1015, "Diameter": "1 inch"},
        {"X": 2676.7327, "Y": 1212.6060, "Lenght": 1985, "Diameter": "1 inch"},
        {"X": 2676.7327, "Y": 1241.6060, "Lenght": 530, "Diameter": "1 inch"},
        {"X": 2684.9600, "Y": 1236.8560, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 15
    [
        {"X": 2739.9581, "Y": 1243.1029, "Lenght": 130, "Diameter": "1 inch"},
        {"X": 2739.9581, "Y": 1257.1029, "Lenght": 125, "Diameter": "1 inch"},
        {"X": 2752.0825, "Y": 1264.1029, "Lenght": 2080, "Diameter": "1 inch"},
        {"X": 2752.0825, "Y": 1232.1029, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 16
    [
        {"X": 2786.7235, "Y": 1216.1029, "Lenght": 130, "Diameter": "1 inch"},
        {"X": 2786.7235, "Y": 1230.1029, "Lenght": 125, "Diameter": "1 inch"},
        {"X": 2798.8478, "Y": 1237.1029, "Lenght": 2080, "Diameter": "1 inch"},
        {"X": 2798.8478, "Y": 1205.1029, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 17
    [
        {"X": 2833.4889, "Y": 1189.1029, "Lenght": 130, "Diameter": "1 inch"},
        {"X": 2833.4889, "Y": 1203.1029, "Lenght": 125, "Diameter": "1 inch"},
        {"X": 2845.6132, "Y": 1210.1029, "Lenght": 2080, "Diameter": "1 inch"},
        {"X": 2845.6132, "Y": 1178.1029, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 18
    [
        {"X": 2880.2542, "Y": 1162.1029, "Lenght": 130, "Diameter": "1 inch"},
        {"X": 2880.2542, "Y": 1176.1029, "Lenght": 125, "Diameter": "1 inch"},
        {"X": 2892.3786, "Y": 1183.1029, "Lenght": 2080, "Diameter": "1 inch"},
        {"X": 2892.3786, "Y": 1151.1029, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 19
    [
        {"X": 2927.0196, "Y": 1135.1029, "Lenght": 130, "Diameter": "1 inch"},
        {"X": 2927.0196, "Y": 1149.1029, "Lenght": 125, "Diameter": "1 inch"},
        {"X": 2939.1440, "Y": 1156.1029, "Lenght": 2080, "Diameter": "1 inch"},
        {"X": 2939.1440, "Y": 1124.1029, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 20
    [
        {"X": 2973.7850, "Y": 1108.1029, "Lenght": 130, "Diameter": "1 inch"},
        {"X": 2973.7850, "Y": 1122.1029, "Lenght": 125, "Diameter": "1 inch"},
        {"X": 2985.9093, "Y": 1129.1029, "Lenght": 2080, "Diameter": "1 inch"},
        {"X": 2985.9093, "Y": 1097.1029, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 21
    [
        {"X": 2845.6132, "Y": 1292.1425, "Lenght": 130, "Diameter": "1 inch"},
        {"X": 2845.6132, "Y": 1306.1425, "Lenght": 125, "Diameter": "1 inch"},
        {"X": 2833.4889, "Y": 1299.1425, "Lenght": 2080, "Diameter": "1 inch"},
        {"X": 2833.4889, "Y": 1267.1425, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 22
    [
        {"X": 2888.9145, "Y": 1267.1425, "Lenght": 130, "Diameter": "1 inch"},
        {"X": 2888.9145, "Y": 1281.1425, "Lenght": 125, "Diameter": "1 inch"},
        {"X": 2876.7901, "Y": 1274.1425, "Lenght": 2080, "Diameter": "1 inch"},
        {"X": 2876.7901, "Y": 1242.1425, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 23
    [
        {"X": 2932.2158, "Y": 1242.1425, "Lenght": 130, "Diameter": "1 inch"},
        {"X": 2932.2158, "Y": 1256.1425, "Lenght": 125, "Diameter": "1 inch"},
        {"X": 2920.0914, "Y": 1249.1425, "Lenght": 2080, "Diameter": "1 inch"},
        {"X": 2920.0914, "Y": 1217.1425, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 24
    [
        {"X": 2975.5170, "Y": 1217.1425, "Lenght": 130, "Diameter": "1 inch"},
        {"X": 2975.5170, "Y": 1231.1425, "Lenght": 125, "Diameter": "1 inch"},
        {"X": 2963.3927, "Y": 1224.1425, "Lenght": 2080, "Diameter": "1 inch"},
        {"X": 2963.3927, "Y": 1192.1425, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 25
    [
        {"X": 3036.6547, "Y": 1181.8447, "Lenght": 130, "Diameter": "1 inch"},
        {"X": 3036.6547, "Y": 1195.8447, "Lenght": 125, "Diameter": "1 inch"},
        {"X": 3024.5303, "Y": 1188.8447, "Lenght": 2080, "Diameter": "1 inch"},
        {"X": 3024.5303, "Y": 1156.8447, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 26
    [
        {"X": 3070.7798, "Y": 1162.1425, "Lenght": 130, "Diameter": "1 inch"},
        {"X": 3070.7798, "Y": 1176.1425, "Lenght": 125, "Diameter": "1 inch"},
        {"X": 3058.6555, "Y": 1169.1425, "Lenght": 2080, "Diameter": "1 inch"},
        {"X": 3058.6555, "Y": 1137.1425, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 27
    [
        {"X": 3117.5452, "Y": 1106.1425, "Lenght": 180, "Diameter": "1 inch"},
        {"X": 3125.7724, "Y": 1101.3925, "Lenght": 1, "Diameter": "1 inch"},
    ],
    # Segmento 28
    [
        {"X": 2804.0440, "Y": 1284.1425, "Lenght": 170, "Diameter": "1 inch"},
        {"X": 2812.2712, "Y": 1279.3925, "Lenght": 1, "Diameter": "1 inch"},
    ],
]

Junctions = [
    # Segmento 1
    {"X": 2146.2196, "Y": 1599.3516, "TAG": "720-PP-100", "Pipe": 152, "Vazao": 4.2, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 11
    {"X": 2693.1872, "Y": 1187.1060, "TAG": "0611-PP-155", "Pipe": 152, "Vazao": 2.4, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 2
    {"X": 2736.4885, "Y": 1162.1060, "TAG": "0611-PP-156", "Pipe": 152, "Vazao": 2.4, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    #{"X": 2219.9440, "Y": 1404.8331, "TAG": "Painel 1", "Pipe": 47, "Vazao": , "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 14
    {"X": 2684.9600, "Y": 1236.8560, "TAG": "0611-PP-109", "Pipe": 152, "Vazao": 5.8, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 13
    {"X": 2698.6842, "Y": 1249.2733, "TAG": "0611-PP-110", "Pipe": 152, "Vazao": 5.8, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 15
    {"X": 2752.0825, "Y": 1232.1029, "TAG": "0611-PP-111", "Pipe": 152, "Vazao": 5.2, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 16
    {"X": 2798.8478, "Y": 1205.1029, "TAG": "0611-PP-112", "Pipe": 152, "Vazao": 5.2, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 17
    {"X": 2845.6132, "Y": 1178.1029, "TAG": "0611-PP-113", "Pipe": 152, "Vazao": 5.8, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 18
    {"X": 2892.3786, "Y": 1151.1029, "TAG": "0611-PP-114", "Pipe": 152, "Vazao": 5.8, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 19
    {"X": 2939.1440, "Y": 1124.1029, "TAG": "0611-PP-115", "Pipe": 152, "Vazao": 4.4, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 20
    {"X": 2985.9093, "Y": 1097.1029, "TAG": "0611-PP-116", "Pipe": 152, "Vazao": 3.0, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 8
    {"X": 3041.3350, "Y": 1051.1029, "TAG": "0611-PP-117", "Pipe": 152, "Vazao": 3.6, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 27
    {"X": 3125.7724, "Y": 1101.3925, "TAG": "0611-PP-108", "Pipe": 152, "Vazao": 3.6, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 26
    {"X": 3058.6555, "Y": 1137.1425, "TAG": "0611-PP-107", "Pipe": 152, "Vazao": 3.0, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 25
    {"X": 3024.5303, "Y": 1156.8447, "TAG": "0611-PP-106", "Pipe": 152, "Vazao": 3.6, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 24
    {"X": 2963.3927, "Y": 1192.1425, "TAG": "0611-PP-105", "Pipe": 152, "Vazao": 5.8, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 23
    {"X": 2920.0914, "Y": 1217.1425, "TAG": "0611-PP-104", "Pipe": 152, "Vazao": 5.2, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 22
    {"X": 2876.7901, "Y": 1242.1425, "TAG": "0611-PP-103", "Pipe": 152, "Vazao": 5.2, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 21
    {"X": 2833.4889, "Y": 1267.1425, "TAG": "0611-PP-102", "Pipe": 152, "Vazao": 5.2, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 28
    {"X": 2812.2712, "Y": 1279.3925, "TAG": "0611-PP-101", "Pipe": 152, "Vazao": 5.8, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 6
    {"X": 2734.0699, "Y": 1358.6801, "TAG": "0611-PP-100", "Pipe": 152, "Vazao": 5.8, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 12
    {"X": 2992.5785, "Y": 1394.4301, "TAG": "0611-PP-153", "Pipe": 152, "Vazao": 2.4, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 3
    {"X": 3060.9945, "Y": 1354.9301, "TAG": "0611-PP-154", "Pipe": 152, "Vazao": 2.4, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 5
    {"X": 2662.6229, "Y": 1543.4301, "TAG": "0611-PP-119", "Pipe": 152, "Vazao": 1.4, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 4
    {"X": 2614.6410, "Y": 1572.2278, "TAG": "0611-PP-152", "Pipe": 152, "Vazao": 2.4, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    # Segmento 10
    {"X": 3408.0967, "Y": 1169.5402, "TAG": "0611-PP-122", "Pipe": 152, "Vazao": 4.4, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
    {"X": 3226.2314, "Y": 1267.5402, "TAG": "0611-PP-118", "Pipe": 152, "Vazao": 4.4, "VazaoUnidade": "Nm3/h", "Temperatura": 25, "TemperaturaUnidade": "deg. C", "Fluido": "Air"},
]

# Find the minimum and maximum X and Y values
min_x = min([coord["X"] for segment in Points for coord in segment])
min_y = min([coord["Y"] for segment in Points for coord in segment])
max_y = max([coord["Y"] for segment in Points for coord in segment])

# Invert the grid, so the AFT grid will be compatible with autocad grid.
Points = InvertGridDirection2D(None, max_y, Points)
Junctions = InvertGridDirection1D(None, max_y, Junctions)

# Shift the coordinato to aproximate the diagram to origin.
min_y = min([coord["Y"] for segment in Points for coord in segment])
max_y = max([coord["Y"] for segment in Points for coord in segment])
Points = ShiftGrid2D(coord = None, offset = {"X": min_x, "Y": min_y}, Points = Points)
Junctions = ShiftGrid1D(coord = None, offset = {"X": min_x, "Y": min_y}, Points = Junctions)

'''
NumberOfSegments = len(Points)

# Construct the Teste string with transformed Points, shifted to the origin
Teste = f"1,Pipe,meters,feet,0,123456789123456789,0,{NumberOfSegments}"

for coord in Points:
    Teste += f",{coord['X'] * 10},{coord['Y'] * 10},0"

Teste += ",None,0,0,0,0,0,None,0,None,0,None,None,0,0,None,0,0,None,None,None,0,06 None,1,1,0,0,None,0,,0,None,0,None,0,None,0,None,0,0,None,1,686.0428,392.237,-999999,-999999,13,5,-1,-1,25,-1,-1,-1,0,0,None,0,0,None,1,0,0,,1,-1,None,-1,-1,-1,0,0,None,0,0,-1,None,None,0,0,0,NA,NA,NA,-1,0,0,0,1,1,%ALL%,4,Percent,1,100,2,100,3,100,0,100,2,years,-9999,-9999,0,0,0,None,-1,0,cm,0,bar,0,1,0,None,0,0,4,0,1,5,-1,-1,-1,-1,-1,10,0,None,1,0,None,1,0,-1,0,6,0,AIR,2810321184,0,None,0,None,None,1,0,None,0,0,None,0,None,False,False,False,0,None"
'''

Pipes = f""
Branchs = f""
NumberOfPoints = 2
NumberOfSegments = 1
PipeNumber = 0
BranchNumber = 0
for segment in Points:
    for i in range(0, len(segment) - 1, 1):
        PipeNumber += 1
        Pipes += f"{PipeNumber},Pipe,mm,meters,0,{NumberOfSegments},{segment[i]['Lenght']},0,0,{NumberOfPoints}"
        Pipes += f",{(segment[i + 0]['X'] + DefaultOffset) * 10},{(segment[i + 0]['Y'] + DefaultOffset) * 10},0"
        Pipes += f",{(segment[i + 1]['X'] + DefaultOffset) * 10},{(segment[i + 1]['Y'] + DefaultOffset) * 10},0"
        Pipes += f",None,0,0,0,0,4.0894,cm,0,None,0,None,Standard,0,0.004572,cm,0,0,Steel - ANSI,{segment[i + 1]['Diameter']},STD (schedule 40),0,06 None,1,1,1,0,None,0,,0,,0,None,0,None,0,None,0,0,None,1,369.5319,342.6998,-999999,-999999,13,5,-1,-1,25,-1,-1,-1,0,0,None,0,0,None,1,0,0,,1,-1,None,-1,-1,-1,0,0,None,0,0,-1,None,None,0,0,0,NA,NA,NA,-1,0,0,0,1,1,%ALL%,4,Percent,1,100,2,100,3,100,0,100,2,years,-9999,-9999,0,0,0,None,-1,0,cm,0,bar,0,1,0,None,0,0,4,0,1,5,-1,-1,-1,-1,-1,10,0,None,1,0,None,1,0,-1,0,6,0,AIR,2810321184,0,None,0,None,None,1,0,None,0,0,None,0,None,False,False,False,0,None\n"
    
    if len(segment) > 2:
        for i in range(1, len(segment) - 1, 1):
            BranchNumber += 1
            NumberOfConnectedPipes = 2
            FirstPipe = i + 0
            SecondPipe = i + 1
            Branchs += f"1,{BranchNumber},6,8,Branch,{NumberOfConnectedPipes},{FirstPipe},-{SecondPipe},{(segment[i + 0]['X'] + DefaultOffset) * 10},{(segment[i + 0]['Y'] + DefaultOffset) * 10},-1,-1,-1,0,0,0,0,None,0,0,0,0,None,1,1,0,None,0,None,0,None,None,0,0,,1,10.32504,-16.20944,-999999,-999999,1,5,-1,-1,25,0,,1,-1,,-1,-1,-1,0,2,-1,3,1,%ALL%,3,Percent,1,100,2,100,3,100,0,2,years,-9999,-9999,4,0,5,0,-999,1,0,None,0,0,None,None,None,0,NA,0,0,0,0,0,-1,-1,0,None,-1,-1,-1,-1,-1,-1,0,None,-1,-1,-1,-1,None,1,0,seconds,None,None,1,0,seconds,None\n"

for junction in Junctions:
    BranchNumber += 1
    NumberOfConnectedPipes = 0
    ShowInWorkspace = 3
    #Branchs += f"3,{BranchNumber},6,8,{junction['TAG']},{NumberOfConnectedPipes},-{junction['Pipe']},{(junction['X'] + DefaultOffset) * 10},{(junction['Y'] + DefaultOffset) * 10},-1,-1,-1,0,0,0,0,None,0,0,0,0,None,1,1,0,None,0,None,0,None,None,0,0,,1,4.169922,-16.59717,-999999,-999999,{ShowInWorkspace},5,-1,-1,25,0,,1,-1,,-1,-1,-1,0,2,-1,3,1,%ALL%,3,Percent,1,100,2,100,3,100,0,2,years,-9999,-9999,4,0,5,0,-999,1,0,None,0,0,None,0,Air,0,NA,0,0,0,0,0,-1,-1,0,None,-1,-1,-1,-1,-1,-1,0,None,-1,-1,-1,-1,None,2,0,seconds,None,None,None,2,0,seconds,None,None,0,0,0,-1,0,None,0,1,0,None,0\n"
    Branchs += f"3,{BranchNumber},6,8,{junction['TAG']},{NumberOfConnectedPipes},{(junction['X'] + DefaultOffset) * 10},{(junction['Y'] + DefaultOffset) * 10},-1,-1,-1,0,0,0,0,None,0,0,0,0,None,1,1,0,None,0,None,0,None,None,0,0,,1,4.169922,-16.59717,-999999,-999999,{ShowInWorkspace},5,-1,-1,25,0,,1,-1,,-1,-1,-1,0,2,-1,3,1,%ALL%,3,Percent,1,100,2,100,3,100,0,2,years,-9999,-9999,4,0,5,0,-999,1,{junction['Vazao']},{junction['VazaoUnidade']},0,{junction['Temperatura']},{junction['TemperaturaUnidade']},0,{junction['Fluido']},0,NA,0,0,0,0,0,-1,-1,0,None,-1,-1,-1,-1,-1,-1,0,None,-1,-1,-1,-1,None,2,0,seconds,None,None,None,2,0,seconds,None,None,0,0,0,-1,0,None,0,1,0,None,0\n"

Pipes = f"[PIPES -- SCENARIO #1 -- Base Scenario]\nNumberOfPipes= {PipeNumber}\n" + Pipes

Branchs = f"[JUNCTIONS -- SCENARIO #1 -- Base Scenario]\nNumberOfJunctions= {BranchNumber}\n" + Branchs

with open(r"C:\Users\dboliveira\Desktop\1.txt", "w") as text_file:
    text_file.write(Branchs)
    text_file.write("\n")
    text_file.write(Pipes)
