# -*- coding: utf-8 -*-
"""
技術分析全攻略 · 個股評分分析系統 (Streamlit 版)
由 stock_analyzer_batch_20260819.html 轉換而成，邏輯與原 HTML/JS 版本一致：
- 朱家泓四維度評分（趨勢／K線／均線／成交量，各25分，共100分）
- 回後買上漲 8 條件核對
- 15 種進場型態確認（含「剛突破」：與前一交易日比較的新鮮突破訊號）
- 批次分析摘要表（可依進場條件／評分／型態確認篩選，關鍵字搜尋）
- Plotly K線＋均線＋布林通道＋成交量＋MACD 圖表
- OpenAI API「AI 智能綜合分析」（金鑰只在瀏覽器/本機端使用，不會上傳儲存）
資料來源：FinMind API
"""

import time
import json
import os
import io
import sqlite3
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import requests
import streamlit as st
from plotly.subplots import make_subplots
import plotly.graph_objects as go

st.set_page_config(page_title="技術分析全攻略 · 個股評分分析系統", page_icon="📊", layout="wide")

FINMIND_BASE = "https://api.finmindtrade.com/api/v4/data"
FINMIND_FALLBACK_BASE = "https://api.finmind.tw/api/latest/data"

# ────────────────────────────────────────────────────────────────
# 預設股票清單（上市 / 上櫃 / 我的清單）
# ────────────────────────────────────────────────────────────────
TWSE_LIST = ['1101', '1101B', '1102', '1103', '1104', '1108', '1109', '1110', '1201', '1203', '1210', '1213', '1215', '1216', '1217', '1218', '1219', '1220', '1225', '1227', '1229', '1231', '1232', '1233', '1234', '1235', '1236', '1256', '1301', '1303', '1304', '1305', '1307', '1308', '1309', '1310', '1312', '1312A', '1313', '1314', '1315', '1316', '1319', '1321', '1323', '1324', '1325', '1326', '1337', '1338', '1339', '1340', '1341', '1342', '1402', '1409', '1410', '1413', '1414', '1416', '1417', '1418', '1419', '1423', '1432', '1434', '1435', '1436', '1437', '1438', '1439', '1440', '1441', '1442', '1443', '1444', '1445', '1446', '1447', '1449', '1451', '1452', '1453', '1454', '1455', '1456', '1457', '1459', '1460', '1463', '1464', '1465', '1466', '1467', '1468', '1470', '1471', '1472', '1473', '1474', '1475', '1476', '1477', '1503', '1504', '1506', '1512', '1513', '1514', '1515', '1516', '1517', '1519', '1521', '1522', '1522A', '1524', '1525', '1526', '1527', '1528', '1529', '1530', '1531', '1532', '1533', '1535', '1536', '1537', '1538', '1539', '1540', '1541', '1558', '1560', '1563', '1568', '1582', '1583', '1587', '1590', '1597', '1598', '1603', '1604', '1605', '1608', '1609', '1611', '1612', '1614', '1615', '1616', '1617', '1618', '1623', '1626', '1702', '1707', '1708', '1709', '1710', '1711', '1712', '1713', '1714', '1717', '1718', '1720', '1721', '1722', '1723', '1725', '1726', '1727', '1730', '1731', '1732', '1733', '1734', '1735', '1736', '1737', '1752', '1760', '1762', '1773', '1776', '1783', '1786', '1789', '1795', '1802', '1805', '1806', '1808', '1809', '1810', '1817', '1903', '1904', '1905', '1906', '1907', '1909', '2002', '2002A', '2006', '2007', '2008', '2009', '2010', '2012', '2013', '2014', '2015', '2017', '2020', '2022', '2023', '2024', '2025', '2027', '2028', '2029', '2030', '2031', '2032', '2033', '2034', '2038', '2049', '2059', '2062', '2069', '2072', '2101', '2102', '2103', '2104', '2105', '2106', '2107', '2108', '2109', '2114', '2115', '2201', '2204', '2206', '2207', '2208', '2211', '2227', '2228', '2231', '2233', '2236', '2239', '2241', '2243', '2247', '2248', '2250', '2254', '2258', '2301', '2302', '2303', '2305', '2308', '2312', '2313', '2314', '2316', '2317', '2321', '2323', '2324', '2327', '2328', '2329', '2330', '2331', '2332', '2337', '2338', '2340', '2342', '2344', '2345', '2347', '2348', '2348A', '2349', '2351', '2352', '2353', '2354', '2355', '2356', '2357', '2359', '2360', '2362', '2363', '2364', '2365', '2367', '2368', '2369', '2371', '2373', '2374', '2375', '2376', '2377', '2379', '2382', '2383', '2385', '2387', '2388', '2390', '2392', '2393', '2395', '2397', '2399', '2401', '2402', '2404', '2405', '2406', '2408', '2409', '2412', '2413', '2414', '2415', '2417', '2419', '2420', '2421', '2423', '2424', '2425', '2426', '2427', '2428', '2429', '2430', '2431', '2432', '2433', '2434', '2436', '2438', '2439', '2440', '2441', '2442', '2444', '2449', '2450', '2451', '2453', '2454', '2455', '2457', '2458', '2459', '2460', '2461', '2462', '2464', '2465', '2466', '2467', '2468', '2471', '2472', '2474', '2476', '2477', '2478', '2480', '2481', '2482', '2483', '2484', '2485', '2486', '2488', '2489', '2491', '2492', '2493', '2495', '2496', '2497', '2498', '2501', '2504', '2505', '2506', '2509', '2511', '2514', '2515', '2516', '2520', '2524', '2527', '2528', '2530', '2534', '2535', '2536', '2537', '2538', '2539', '2540', '2542', '2543', '2545', '2546', '2547', '2548', '2597', '2601', '2603', '2605', '2606', '2607', '2608', '2609', '2610', '2611', '2612', '2613', '2614', '2615', '2616', '2617', '2618', '2630', '2633', '2634', '2636', '2637', '2642', '2645', '2646', '2701', '2702', '2704', '2705', '2706', '2707', '2712', '2722', '2723', '2727', '2731', '2739', '2748', '2753', '2762', '2801', '2812', '2816', '2820', '2832', '2834', '2836', '2836A', '2838', '2838A', '2845', '2849', '2850', '2851', '2852', '2855', '2867', '2880', '2881', '2881A', '2881B', '2881C', '2882', '2882A', '2882B', '2883', '2883B', '2884', '2885', '2886', '2887', '2887E', '2887F', '2887G', '2887H', '2887I', '2887Z1', '2889', '2890', '2891', '2891B', '2891C', '2892', '2897', '2897B', '2901', '2903', '2904', '2905', '2906', '2908', '2910', '2911', '2912', '2913', '2915', '2923', '2929', '2939', '2945', '3002', '3003', '3004', '3005', '3006', '3008', '3010', '3011', '3013', '3014', '3015', '3016', '3017', '3018', '3019', '3021', '3022', '3023', '3024', '3025', '3026', '3027', '3028', '3029', '3030', '3031', '3032', '3033', '3034', '3035', '3036', '3037', '3038', '3040', '3041', '3042', '3043', '3044', '3045', '3046', '3047', '3048', '3049', '3050', '3051', '3052', '3054', '3055', '3056', '3057', '3058', '3059', '3060', '3062', '3090', '3092', '3094', '3130', '3135', '3138', '3149', '3150', '3164', '3167', '3168', '3189', '3209', '3229', '3231', '3257', '3266', '3296', '3305', '3308', '3311', '3312', '3321', '3338', '3346', '3356', '3376', '3380', '3406', '3413', '3416', '3419', '3432', '3437', '3443', '3447', '3450', '3481', '3494', '3501', '3504', '3515', '3518', '3528', '3530', '3532', '3533', '3535', '3543', '3545', '3550', '3557', '3563', '3576', '3583', '3588', '3591', '3592', '3593', '3596', '3605', '3607', '3617', '3622', '3645', '3652', '3653', '3661', '3665', '3669', '3673', '3679', '3686', '3694', '3701', '3702', '3703', '3704', '3705', '3706', '3708', '3711', '3712', '3714', '3715', '3716', '3717', '4104', '4106', '4108', '4119', '4133', '4137', '4142', '4148', '4155', '4164', '4169', '4178', '4190', '4195', '4306', '4414', '4426', '4438', '4439', '4440', '4441', '4526', '4532', '4536', '4540', '4545', '4551', '4552', '4555', '4557', '4560', '4562', '4564', '4566', '4569', '4571', '4572', '4576', '4581', '4582', '4583', '4585', '4588', '4590', '4720', '4722', '4736', '4737', '4739', '4746', '4755', '4763', '4764', '4766', '4770', '4771', '4807', '4904', '4906', '4912', '4915', '4916', '4919', '4927', '4930', '4934', '4935', '4938', '4942', '4943', '4949', '4952', '4956', '4958', '4960', '4961', '4967', '4968', '4976', '4977', '4989', '4994', '4999', '5007', '5203', '5215', '5222', '5225', '5234', '5243', '5244', '5258', '5269', '5283', '5284', '5285', '5288', '5292', '5306', '5388', '5434', '5469', '5471', '5484', '5515', '5519', '5521', '5522', '5525', '5531', '5533', '5534', '5538', '5546', '5607', '5608', '5706', '5871', '5871A', '5876', '5880', '5906', '5907', '6005', '6024', '6108', '6112', '6115', '6116', '6117', '6120', '6128', '6133', '6136', '6139', '6141', '6142', '6152', '6153', '6155', '6164', '6165', '6166', '6168', '6176', '6177', '6183', '6184', '6189', '6191', '6192', '6196', '6197', '6201', '6202', '6205', '6206', '6209', '6213', '6214', '6215', '6216', '6224', '6225', '6226', '6230', '6235', '6239', '6243', '6257', '6269', '6271', '6272', '6277', '6278', '6281', '6282', '6283', '6285', '6405', '6409', '6412', '6414', '6415', '6416', '6426', '6431', '6438', '6442', '6443', '6446', '6449', '6451', '6456', '6464', '6472', '6477', '6491', '6504', '6505', '6515', '6525', '6526', '6531', '6533', '6534', '6541', '6550', '6552', '6558', '6573', '6579', '6581', '6582', '6585', '6589', '6591', '6592', '6592A', '6592B', '6598', '6605', '6606', '6614', '6625', '6641', '6645', '6655', '6657', '6658', '6666', '6668', '6669', '6670', '6671', '6672', '6674', '6689', '6691', '6695', '6698', '6706', '6715', '6719', '6722', '6742', '6743', '6753', '6754', '6756', '6757', '6768', '6770', '6771', '6776', '6781', '6782', '6789', '6790', '6792', '6794', '6796', '6799', '6805', '6807', '6830', '6831', '6834', '6835', '6838', '6854', '6861', '6862', '6863', '6869', '6873', '6885', '6887', '6890', '6901', '6902', '6906', '6908', '6909', '6914', '6916', '6918', '6919', '6921', '6923', '6924', '6928', '6931', '6933', '6934', '6936', '6937', '6944', '6949', '6951', '6952', '6955', '6957', '6958', '6958A', '6962', '6965', '6969', '6988', '6994', '7610', '7631', '7705', '7711', '7721', '7722', '7730', '7732', '7736', '7740', '7749', '7750', '7760', '7765', '7768', '7769', '7780', '7786', '7788', '7791', '7795', '7799', '7803', '7818', '7821', '7822', '7823', '7827', '8011', '8016', '8021', '8028', '8033', '8039', '8045', '8046', '8070', '8072', '8081', '8101', '8103', '8104', '8105', '8110', '8112', '8112A', '8114', '8131', '8150', '8162', '8163', '8201', '8210', '8213', '8215', '8222', '8249', '8261', '8271', '8341', '8367', '8374', '8404', '8411', '8422', '8429', '8438', '8442', '8443', '8454', '8462', '8463', '8464', '8466', '8467', '8473', '8476', '8478', '8481', '8482', '8487', '8488', '8499', '8926', '8940', '8996', '9103', '910322', '9105', '910861', '9110', '911608', '911622', '911868', '912000', '9136', '9802', '9902', '9904', '9905', '9906', '9907', '9908', '9910', '9911', '9912', '9914', '9917', '9918', '9919', '9921', '9924', '9925', '9926', '9927', '9928', '9929', '9930', '9931', '9933', '9934', '9935', '9937', '9938', '9939', '9940', '9941', '9941A', '9942', '9943', '9944', '9945', '9946', '9955', '9958']

TPEX_LIST = ['1240', '1259', '1264', '1268', '1294', '1295', '1336', '1565', '1569', '1570', '1580', '1584', '1586', '1591', '1593', '1595', '1599', '1742', '1777', '1780', '1781', '1784', '1785', '1788', '1796', '1799', '1813', '1815', '2035', '2061', '2063', '2064', '2065', '2066', '2067', '2070', '2073', '2221', '2230', '2235', '2596', '2640', '2641', '2643', '2718', '2719', '2724', '2726', '2729', '2732', '2734', '2736', '2740', '2743', '2745', '2751', '2752', '2754', '2755', '2756', '2916', '2924', '2926', '2937', '2941', '2947', '2948', '2949', '3064', '3066', '3067', '3071', '3073', '3078', '3081', '3083', '3085', '3086', '3088', '3093', '3095', '3105', '3114', '3115', '3118', '3122', '3128', '3131', '3141', '3147', '3152', '3158', '3162', '3163', '3169', '3171', '3176', '3178', '3188', '3191', '3205', '3206', '3207', '3211', '3213', '3217', '3218', '3219', '3221', '3224', '3226', '3227', '3228', '3230', '3232', '3234', '3236', '3252', '3259', '3260', '3264', '3265', '3268', '3272', '3276', '3284', '3285', '3287', '3288', '3289', '3290', '3293', '3294', '3297', '3303', '3306', '3310', '3313', '3317', '3322', '3323', '3324', '3325', '3332', '3339', '3349', '3354', '3357', '3360', '3362', '3363', '3372', '3373', '3374', '3379', '3388', '3390', '3402', '3430', '3434', '3438', '3441', '3444', '3455', '3465', '3466', '3467', '3479', '3483', '3484', '3485', '3489', '3490', '3491', '3492', '3498', '3499', '3508', '3511', '3512', '3516', '3520', '3521', '3522', '3523', '3526', '3527', '3529', '3531', '3537', '3540', '3541', '3546', '3548', '3551', '3552', '3555', '3556', '3558', '3564', '3567', '3570', '3577', '3580', '3581', '3587', '3594', '3597', '3609', '3611', '3615', '3623', '3624', '3625', '3628', '3629', '3630', '3631', '3632', '3646', '3663', '3664', '3666', '3672', '3675', '3680', '3684', '3685', '3687', '3689', '3691', '3693', '3707', '3709', '3710', '3713', '4102', '4105', '4107', '4109', '4111', '4113', '4114', '4116', '4120', '4121', '4123', '4126', '4127', '4128', '4129', '4131', '4138', '4139', '4147', '4153', '4154', '4157', '4160', '4161', '4162', '4163', '4166', '4167', '4168', '4171', '4173', '4174', '4175', '4183', '4188', '4192', '4198', '4205', '4207', '4303', '4304', '4305', '4401', '4402', '4406', '4413', '4416', '4417', '4419', '4420', '4430', '4432', '4433', '4442', '4502', '4503', '4506', '4510', '4513', '4523', '4527', '4528', '4529', '4530', '4533', '4534', '4535', '4538', '4541', '4542', '4543', '4549', '4550', '4554', '4556', '4558', '4561', '4563', '4568', '4577', '4580', '4584', '4609', '4702', '4706', '4707', '4711', '4714', '4716', '4721', '4726', '4728', '4729', '4735', '4741', '4743', '4744', '4745', '4747', '4749', '4754', '4760', '4767', '4768', '4772', '4806', '4903', '4905', '4907', '4908', '4909', '4911', '4923', '4924', '4931', '4933', '4939', '4946', '4950', '4951', '4953', '4966', '4971', '4972', '4973', '4974', '4979', '4991', '4995', '5009', '5011', '5013', '5014', '5015', '5016', '5201', '5202', '5205', '5206', '5209', '5210', '5211', '5212', '5213', '5220', '5223', '5227', '5228', '5230', '5245', '5251', '5263', '5272', '5274', '5276', '5278', '5287', '5289', '5291', '5299', '5301', '5302', '5309', '5310', '5312', '5314', '5315', '5321', '5324', '5328', '5340', '5344', '5345', '5347', '5348', '5351', '5353', '5355', '5356', '5364', '5371', '5381', '5386', '5392', '5398', '5403', '5410', '5425', '5426', '5432', '5438', '5439', '5443', '5450', '5452', '5455', '5457', '5460', '5464', '5465', '5468', '5474', '5475', '5478', '5481', '5483', '5487', '5488', '5489', '5490', '5493', '5498', '5508', '5511', '5512', '5514', '5516', '5520', '5523', '5529', '5530', '5536', '5543', '5547', '5548', '5601', '5603', '5604', '5609', '5701', '5703', '5704', '5864', '5878', '5902', '5903', '5904', '5905', '6015', '6016', '6020', '6021', '6023', '6026', '6028', '6101', '6103', '6104', '6109', '6111', '6113', '6114', '6118', '6121', '6122', '6123', '6124', '6125', '6126', '6127', '6129', '6130', '6134', '6138', '6140', '6143', '6144', '6146', '6147', '6148', '6150', '6151', '6154', '6156', '6158', '6160', '6161', '6163', '6167', '6169', '6170', '6171', '6173', '6174', '6175', '6179', '6180', '6182', '6185', '6186', '6187', '6188', '6190', '6194', '6195', '6198', '6199', '6203', '6204', '6207', '6208', '6210', '6212', '6217', '6218', '6219', '6220', '6221', '6222', '6223', '6227', '6228', '6229', '6231', '6233', '6234', '6236', '6237', '6240', '6241', '6242', '6244', '6245', '6246', '6248', '6259', '6261', '6263', '6264', '6265', '6266', '6270', '6274', '6275', '6276', '6279', '6284', '6290', '6291', '6292', '6294', '6411', '6417', '6418', '6419', '6423', '6425', '6432', '6435', '6441', '6461', '6462', '6465', '6469', '6470', '6474', '6482', '6485', '6486', '6488', '6492', '6494', '6496', '6498', '6499', '6506', '6508', '6509', '6510', '6512', '6516', '6517', '6523', '6527', '6530', '6532', '6535', '6538', '6542', '6546', '6547', '6548', '6556', '6560', '6561', '6568', '6569', '6570', '6574', '6576', '6577', '6578', '6584', '6588', '6590', '6593', '6596', '6597', '6603', '6609', '6612', '6613', '6615', '6616', '6617', '6620', '6624', '6629', '6637', '6640', '6642', '6643', '6649', '6651', '6654', '6661', '6662', '6664', '6667', '6679', '6680', '6683', '6684', '6690', '6692', '6693', '6697', '6703', '6708', '6712', '6716', '6720', '6721', '6725', '6727', '6728', '6730', '6732', '6733', '6735', '6739', '6741', '6751', '6752', '6761', '6762', '6763', '6767', '6785', '6788', '6791', '6803', '6804', '6811', '6821', '6823', '6829', '6840', '6841', '6843', '6844', '6846', '6855', '6856', '6859', '6865', '6870', '6872', '6874', '6875', '6877', '6881', '6884', '6894', '6895', '6899', '6903', '6904', '6907', '6910', '6913', '6922', '6925', '6929', '6945', '6953', '6961', '6967', '6968', '6971', '6982', '6983', '6986', '6996', '6997', '7402', '7547', '7556', '7584', '7642', '7703', '7704', '7708', '7709', '7712', '7713', '7714', '7715', '7716', '7717', '7718', '7723', '7728', '7734', '7738', '7743', '7744', '7747', '7751', '7753', '7757', '7767', '7770', '7772', '7777', '7782', '7792', '7794', '7805', '7810', '7811', '7814', '7819', '7820', '7828', '7839', '7842', '8024', '8027', '8032', '8034', '8038', '8040', '8042', '8043', '8044', '8047', '8048', '8049', '8050', '8054', '8059', '8064', '8066', '8067', '8068', '8069', '8071', '8074', '8076', '8077', '8080', '8083', '8084', '8085', '8086', '8087', '8088', '8089', '8091', '8092', '8093', '8096', '8097', '8099', '8102', '8107', '8109', '8111', '8121', '8147', '8155', '8171', '8176', '8182', '8183', '8227', '8234', '8240', '8255', '8272', '8277', '8279', '8284', '8289', '8291', '8299', '8342', '8349', '8349A', '8354', '8358', '8383', '8390', '8401', '8403', '8409', '8410', '8415', '8416', '8421', '8423', '8424', '8426', '8431', '8432', '8433', '8435', '8436', '8437', '8440', '8444', '8446', '8450', '8455', '8472', '8477', '8489', '8905', '8906', '8908', '8916', '8917', '8921', '8923', '8924', '8927', '8928', '8929', '8930', '8931', '8932', '8933', '8935', '8936', '8937', '8938', '8941', '8942', '9949', '9950', '9951', '9960', '9962']

MY_LIST = ['4979', '6870', '6451', '3450', '3163', '4977', '3363', '6533', '6757', '2345', '2376', '2441', '6285', '2049', '2303', '3653', '2467', '2359', '2308', '3017', '2368', '4576', '2327', '6442', '6531', '6683', '6257', '6223', '8150', '3105', '2634', '8299', '3532', '6213', '3033', '4991', '2449', '6781', '3189', '8046', '3037', '3443', '3491', '2481', '6770', '2408', '8271', '3081', '3665', '6274', '2455', '2454', '2360', '3673', '8021', '4573', '2351', '1560', '8028', '6197', '4958', '8358', '2330', '2383', '1303', '6669', '2137', '3231', '1815', '5475', '5340', '3771', '1519', '3661', '2059', '6805', '6510', '3211', '4931', '8210', '3013', '6117', '3693']

# 0050（元大台灣50）成分股：追蹤「臺灣50指數」，每季（3/6/9/12月）審核調整一次，
# 是規則化的被動指數，成分股清單相對穩定、可公開查證，適合寫死維護。
# 以下為 2026-09-04 公開持股快照（共51檔），下次調整約在2026年12月，屆時建議
# 重新核對更新。注意：00981A／00991A是主動式ETF（經理人自由調整持股、無固定
# 規則、且無完整成分股API），沒有一併做在這裡——硬編碼一份很快就會失準。
TW0050_LIST = ['2330', '2454', '2308', '2317', '3711', '2383', '2303', '2881', '2891', '3037', '3017', '1303', '2882', '2345', '2887', '2382', '2327', '2885', '2357', '2884', '3008', '2301', '3231', '2886', '2883', '2408', '2890', '2344', '2412', '2892', '2449', '2337', '4938', '2356', '2368', '5880', '1301', '1326', '2002', '1101', '1216', '3045', '4904', '2379', '3034', '1102', '2610', '2618', '2603', '2609', '2615']


# ────────────────────────────────────────────────────────────────
# 自訂清單（我的清單／清單1／清單2／清單3）：存成本機JSON檔，跨次啟動App都會保留
# （對應 HTML 版用 localStorage 的效果，只是這裡改用本機檔案，因為 Streamlit
# 的 session_state 每次重新啟動就會清空，沒有等同 localStorage 的內建機制）。
# 「我的清單」預設帶入原本內建的 MY_LIST 觀察名單，清單1/2/3預設空白。
# ────────────────────────────────────────────────────────────────
CUSTOM_LISTS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "my_lists.json")
CUSTOM_LIST_KEYS = ["my", "my1", "my2", "my3"]
CUSTOM_LIST_LABELS = {"my": "我的清單", "my1": "我的清單1", "my2": "我的清單2", "my3": "我的清單3"}


def load_custom_lists():
    defaults = {"my": list(MY_LIST), "my1": [], "my2": [], "my3": []}
    if os.path.exists(CUSTOM_LISTS_FILE):
        try:
            with open(CUSTOM_LISTS_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            for k in CUSTOM_LIST_KEYS:
                if k in saved and isinstance(saved[k], list):
                    defaults[k] = saved[k]
        except Exception:
            pass
    return defaults


def save_custom_lists(lists: dict):
    try:
        with open(CUSTOM_LISTS_FILE, "w", encoding="utf-8") as f:
            json.dump(lists, f, ensure_ascii=False)
    except Exception:
        pass


def parse_stock_tokens(text: str):
    stocks, seen = [], set()
    for tok in text.replace("，", ",").replace("、", ",").split():
        for s in tok.split(","):
            s = s.strip()
            if s and s not in seen:
                seen.add(s)
                stocks.append(s)
    return stocks


# ────────────────────────────────────────────────────────────────
# FinMind API 存取（含備援 endpoint，與原 HTML 版一致）
# ────────────────────────────────────────────────────────────────

def get_date_range(days_back: int):
    end = datetime.today()
    start = datetime.today() - timedelta(days=days_back)
    return start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")


def api_fetch(url_v4: str):
    """依序嘗試 FinMind v4 endpoint 及備援 endpoint。"""
    url_fallback = url_v4.replace("api.finmindtrade.com/api/v4", "api.finmind.tw/api/latest")
    last_err = ""
    for url in (url_v4, url_fallback):
        try:
            r = requests.get(url, timeout=15)
            if not r.ok:
                last_err = f"HTTP {r.status_code}"
                continue
            j = r.json()
            if j.get("status") == 200:
                return j
            last_err = j.get("msg") or f"status={j.get('status')}"
        except Exception as e:
            last_err = str(e)
    raise RuntimeError(last_err or "fetch failed")


def fetch_price_data(stock_id: str, token: str, days: int):
    start, end = get_date_range(days)
    url = f"{FINMIND_BASE}?dataset=TaiwanStockPrice&data_id={stock_id}&start_date={start}&end_date={end}&token={token}"
    j = api_fetch(url)
    if not j.get("data"):
        raise RuntimeError("無資料（可能代號錯誤）")
    return j["data"]


# ── 近3年 P/E 區間（TaiwanStockPER，每股票每次批次分析都會多打一次API）──
# ETF（如0050）本身沒有EPS，FinMind不會回傳PER資料，這種情況回傳None，不是bug。
# 取得的PER陣列先過濾掉0、負值、非數字（虧損股常見PER<=0或缺值），再取這3年內
# 的最大最小值作為區間，最新一筆視為目前P/E。
def fetch_pe_range(stock_id: str, token: str, years: int = 3):
    end = datetime.today()
    start = datetime.today() - timedelta(days=365 * years)
    fmt = "%Y-%m-%d"
    url = f"{FINMIND_BASE}?dataset=TaiwanStockPER&data_id={stock_id}&start_date={start.strftime(fmt)}&end_date={end.strftime(fmt)}&token={token}"
    j = api_fetch(url)
    rows = j.get("data") or []
    if not rows:
        return None
    rows = sorted(rows, key=lambda r: r["date"])
    values = []
    for r in rows:
        try:
            v = float(r.get("PER"))
            if v > 0 and v == v:  # v==v 排除 NaN
                values.append(v)
        except (TypeError, ValueError):
            continue
    if not values:
        return None
    try:
        current = float(rows[-1].get("PER"))
        if not (current > 0):
            current = values[-1]
    except (TypeError, ValueError):
        current = values[-1]
    return {"min": min(values), "max": max(values), "current": current}


# ── 近3個月營收 YoY／MoM（TaiwanStockMonthRevenue，每股票每次批次分析都會多打一次API）──
# YoY／MoM 是自己算的（FinMind只給原始revenue，不含增減率），需要抓夠長的區間
# （往前16個月）才能算出最近3個月當中「最舊那個月」的YoY（要對比去年同月）。
def fetch_revenue_yoy_mom(stock_id: str, token: str):
    end = datetime.today()
    start = datetime.today() - timedelta(days=30 * 16)
    fmt = "%Y-%m-%d"
    url = f"{FINMIND_BASE}?dataset=TaiwanStockMonthRevenue&data_id={stock_id}&start_date={start.strftime(fmt)}&end_date={end.strftime(fmt)}&token={token}"
    j = api_fetch(url)
    rows = j.get("data") or []
    if not rows:
        return None
    rev_map = {}
    for r in rows:
        key = f"{r['revenue_year']}-{r['revenue_month']:02d}"
        rev_map[key] = r["revenue"]
    months = sorted(rev_map.keys())
    last3 = months[-3:]
    if not last3:
        return None
    result = []
    for key in last3:
        y, m = (int(p) for p in key.split("-"))
        rev = rev_map[key]
        prev_y, prev_m = (y - 1, 12) if m == 1 else (y, m - 1)
        prev_key = f"{prev_y}-{prev_m:02d}"
        yoy_key = f"{y - 1}-{m:02d}"
        mom = ((rev - rev_map[prev_key]) / rev_map[prev_key] * 100) if rev_map.get(prev_key) else None
        yoy = ((rev - rev_map[yoy_key]) / rev_map[yoy_key] * 100) if rev_map.get(yoy_key) else None
        result.append({"year": y, "month": m, "revenue": rev, "mom": mom, "yoy": yoy})
    return result


# ── 三大法人買賣超（近3個月，逐月加總）──
# TaiwanStockInstitutionalInvestorsBuySell是「每日」資料，不是月資料，這裡自己依日期
# 分組加總成月度買賣超。抓約100天回來，涵蓋最近3個完整月份綽綽有餘。
# 買賣超 = buy - sell（單日、單一法人類別），三大法人合計 = 外資+投信+自營商
# 三類加總。正值=買超（淨買進），負值=賣超（淨賣出）。單位是FinMind原始股數，
# 沒有換算成「張」。
def fetch_institutional_monthly(stock_id: str, token: str):
    end = datetime.today()
    start = datetime.today() - timedelta(days=100)
    fmt = "%Y-%m-%d"
    url = f"{FINMIND_BASE}?dataset=TaiwanStockInstitutionalInvestorsBuySell&data_id={stock_id}&start_date={start.strftime(fmt)}&end_date={end.strftime(fmt)}&token={token}"
    j = api_fetch(url)
    rows = j.get("data") or []
    if not rows:
        return None
    month_map = {}
    for r in rows:
        d = str(r.get("date", ""))[:10]
        if len(d) < 7:
            continue
        key = d[:7]  # "YYYY-MM"
        net = (float(r.get("buy") or 0)) - (float(r.get("sell") or 0))
        month_map[key] = month_map.get(key, 0) + net
    months = sorted(month_map.keys())
    last3 = months[-3:]
    if not last3:
        return None
    result = []
    for key in last3:
        y, m = (int(p) for p in key.split("-"))
        result.append({"year": y, "month": m, "net": month_map[key]})
    return result


def _last_n_calendar_months(n: int):
    out = []
    d = datetime.today().replace(day=1)
    for _ in range(n):
        out.insert(0, {"year": d.year, "month": d.month})
        d = (d - timedelta(days=1)).replace(day=1)
    return out


# ── 近3個月「當月均價」YoY（跟營收YoY用同一組月份，方便橫向比較；若沒給月份清單
# 則自己算「最近3個日曆月」，讓這個功能可以獨立於營收YoY單獨使用）──
# 自己抓一段股價區間、按年月分組算平均收盤價，再跟去年同月的均價比。這是獨立於
# 「分析天數」的另一次 TaiwanStockPrice 查詢（往前抓到最舊那個月的去年同月，
# 通常要抓超過一年），每股票每次批次分析又會多打一次API。
def fetch_monthly_avg_price_yoy(stock_id: str, token: str, months_list=None):
    if not months_list:
        months_list = _last_n_calendar_months(3)
    oldest = months_list[0]
    start = datetime(oldest["year"] - 1, oldest["month"], 1)
    end = datetime.today()
    fmt = "%Y-%m-%d"
    url = f"{FINMIND_BASE}?dataset=TaiwanStockPrice&data_id={stock_id}&start_date={start.strftime(fmt)}&end_date={end.strftime(fmt)}&token={token}"
    j = api_fetch(url)
    rows = j.get("data") or []
    if not rows:
        return None
    sums, counts = {}, {}
    for r in rows:
        key = str(r["date"])[:7]
        try:
            c = float(r["close"])
        except (TypeError, ValueError):
            continue
        if c <= 0:
            continue
        sums[key] = sums.get(key, 0) + c
        counts[key] = counts.get(key, 0) + 1
    avg_map = {k: sums[k] / counts[k] for k in sums}
    result = []
    for m in months_list:
        key = f"{m['year']}-{m['month']:02d}"
        yoy_key = f"{m['year'] - 1}-{m['month']:02d}"
        this_avg, last_avg = avg_map.get(key), avg_map.get(yoy_key)
        yoy = ((this_avg - last_avg) / last_avg * 100) if (this_avg is not None and last_avg) else None
        result.append({"year": m["year"], "month": m["month"], "avg": this_avg, "avgLastYear": last_avg, "yoy": yoy})
    return result


# ── 盤中即時股價快照（taiwan_stock_tick_snapshot，約10秒更新一次）──
# 注意：這是 FinMind「只限sponsor會員使用」的端點，跟一般 dataset= 資料不同路徑
# （https://api.finmindtrade.com/api/v4/taiwan_stock_tick_snapshot，不是 .../data?dataset=...）。
# 非sponsor會員呼叫會被拒絕，這裡設計成失敗就整批跳過、不中斷主流程。
REALTIME_SNAPSHOT_URL = "https://api.finmindtrade.com/api/v4/taiwan_stock_tick_snapshot"
REALTIME_CHUNK_SIZE = 80


def _chunk_list(lst, size):
    return [lst[i:i + size] for i in range(0, len(lst), size)]


def fetch_realtime_snapshots(token: str, stock_ids: list):
    """回傳 (snapshot_map, error_msg)。snapshot_map: {stock_id: row}；
    error_msg 有值代表整批都失敗（例如非sponsor會員），此時 snapshot_map 為空。"""
    snapshot_map = {}
    last_err = None
    for chunk in _chunk_list(stock_ids, REALTIME_CHUNK_SIZE):
        params = "&".join(f"data_id={sid}" for sid in chunk)
        url = f"{REALTIME_SNAPSHOT_URL}?{params}&token={token}"
        try:
            r = requests.get(url, timeout=15)
            if not r.ok:
                last_err = f"HTTP {r.status_code}：{r.text[:200]}"
                continue
            j = r.json()
            if j.get("status") != 200:
                last_err = j.get("msg") or f"status={j.get('status')}"
                continue
            for row in j.get("data", []):
                sid = row.get("stock_id")
                if sid:
                    snapshot_map[sid] = row
        except Exception as e:
            last_err = str(e)
    return snapshot_map, (None if snapshot_map else last_err)


def merge_realtime_snapshot(raw_data: list, snapshot: dict):
    """把即時快照併入歷史資料的最後一筆：只有在「今天還沒被包含在EOD資料裡」時才附加
    一筆合成的當日K棒，讓評分／型態辨識／圖表在盤中也能反映當下股價。
    回傳 (data, injected)。"""
    if not snapshot or not snapshot.get("date"):
        return raw_data, False
    today_str = str(snapshot["date"])[:10]
    last_date = raw_data[-1]["date"] if raw_data else None
    if last_date and today_str <= last_date:
        return raw_data, False
    try:
        close = float(snapshot.get("close") or 0)
    except (TypeError, ValueError):
        close = 0
    if not close:
        return raw_data, False
    try:
        open_ = float(snapshot.get("open") or close)
        high = float(snapshot.get("high") or close)
        low = float(snapshot.get("low") or close)
    except (TypeError, ValueError):
        open_, high, low = close, close, close
    vol_raw = snapshot.get("total_volume", snapshot.get("volume"))
    try:
        vol = float(vol_raw) if vol_raw is not None else 0
    except (TypeError, ValueError):
        vol = 0
    merged = raw_data + [{
        "date": today_str, "open": open_, "high": high, "low": low, "close": close, "volume": vol,
    }]
    return merged, True


@st.cache_data(show_spinner=False, ttl=3600)
def fetch_stock_name_map(token: str) -> dict:
    try:
        url = f"{FINMIND_BASE}?dataset=TaiwanStockInfo&token={token}"
        j = api_fetch(url)
        name_map = {}
        for item in j.get("data", []):
            sid = item.get("stock_id")
            sname = item.get("stock_name")
            if sid and sname and sid not in name_map:
                name_map[sid] = sname
        return name_map
    except Exception:
        return {}


# ────────────────────────────────────────────────────────────────
# 技術指標計算（與原 HTML/JS 版本邏輯一致）
# ────────────────────────────────────────────────────────────────

def calc_ma(closes, p):
    out = []
    for i in range(len(closes)):
        if i < p - 1:
            out.append(None)
        else:
            out.append(sum(closes[i - p + 1:i + 1]) / p)
    return out


def calc_ema_series(closes, p):
    k = 2 / (p + 1)
    e = closes[0]
    out = [e]
    for i in range(1, len(closes)):
        e = closes[i] * k + e * (1 - k)
        out.append(e)
    return out


def calc_macd_series(closes):
    e12 = calc_ema_series(closes, 12)
    e26 = calc_ema_series(closes, 26)
    dif = [a - b for a, b in zip(e12, e26)]
    k = 2 / 10
    s = dif[0]
    sig = [s]
    for i in range(1, len(dif)):
        s = dif[i] * k + s * (1 - k)
        sig.append(s)
    hist = [d - sgl for d, sgl in zip(dif, sig)]
    return {"dif": dif, "sig": sig, "hist": hist}


def calc_rsi_series(closes, p=14):
    n = len(closes)
    res = [None] * n
    ag = 0.0
    al = 0.0
    for i in range(1, p + 1):
        d = closes[i] - closes[i - 1]
        if d > 0:
            ag += d
        else:
            al += abs(d)
    ag /= p
    al /= p
    res[p] = 100 - 100 / (1 + ag / (al or 0.001))
    for i in range(p + 1, n):
        d = closes[i] - closes[i - 1]
        ag = (ag * (p - 1) + (d if d > 0 else 0)) / p
        al = (al * (p - 1) + (abs(d) if d < 0 else 0)) / p
        res[i] = 100 - 100 / (1 + ag / (al or 0.001))
    return res


def calc_bb_series(closes, p=20, std=2):
    mid = calc_ma(closes, p)
    out = []
    for i in range(len(closes)):
        if i < p - 1:
            out.append({"u": None, "l": None})
            continue
        m = mid[i]
        window = closes[i - p + 1:i + 1]
        s = sum((c - m) ** 2 for c in window)
        sigma = (s / p) ** 0.5
        out.append({"u": m + std * sigma, "l": m - std * sigma})
    return out


def calc_volma(volumes, p):
    out = []
    for i in range(len(volumes)):
        if i < p - 1:
            out.append(None)
        else:
            out.append(sum(volumes[i - p + 1:i + 1]) / p)
    return out


def pivots(data, w=5):
    highs, lows = [], []
    n = len(data)
    for i in range(w, n - w):
        h, l = data[i]["high"], data[i]["low"]
        is_h, is_l = True, True
        for j in range(i - w, i + w + 1):
            if data[j]["high"] > h:
                is_h = False
            if data[j]["low"] < l:
                is_l = False
        if is_h:
            highs.append(i)
        if is_l:
            lows.append(i)
    return {"highs": highs, "lows": lows}


def calc_kd(data, period=9):
    k_arr, d_arr = [], []
    prev_k, prev_d = 50.0, 50.0
    for i in range(len(data)):
        if i < period - 1:
            k_arr.append(None)
            d_arr.append(None)
            continue
        window = data[i - period + 1:i + 1]
        lowest = min(d["low"] for d in window)
        highest = max(d["high"] for d in window)
        rsv = 50 if highest == lowest else (data[i]["close"] - lowest) / (highest - lowest) * 100
        k = prev_k * 2 / 3 + rsv * 1 / 3
        d = prev_d * 2 / 3 + k * 1 / 3
        k_arr.append(k)
        d_arr.append(d)
        prev_k, prev_d = k, d
    return {"k": k_arr, "d": d_arr}


def enrich(data):
    closes = [d["close"] for d in data]
    volumes = [d["volume"] for d in data]
    m5, m10, m20, m60 = calc_ma(closes, 5), calc_ma(closes, 10), calc_ma(closes, 20), calc_ma(closes, 60)
    md = calc_macd_series(closes)
    rs = calc_rsi_series(closes)
    bbs = calc_bb_series(closes)
    vm5, vm20 = calc_volma(volumes, 5), calc_volma(volumes, 20)
    kd = calc_kd(data, 9)
    out = []
    for i, d in enumerate(data):
        out.append({
            "date": d["date"], "open": d["open"], "high": d["high"], "low": d["low"],
            "close": d["close"], "volume": d["volume"],
            "ma5": m5[i], "ma10": m10[i], "ma20": m20[i], "ma60": m60[i],
            "macd": md["dif"][i], "macdSig": md["sig"][i], "macdHist": md["hist"][i],
            "rsi": rs[i], "bbU": bbs[i]["u"], "bbL": bbs[i]["l"],
            "vm5": vm5[i], "vm20": vm20[i],
            "kdK": kd["k"][i], "kdD": kd["d"][i],
        })
    return out


# ────────────────────────────────────────────────────────────────
# 朱家泓四維度評分（趨勢／K線／均線／成交量，各25分）
# ────────────────────────────────────────────────────────────────

def score_trend(data):
    score, sigs = 0, []
    last = data[-1]
    tdir = "盤整"
    pv = pivots(data)
    if len(pv["highs"]) >= 2 and len(pv["lows"]) >= 2:
        rh = [data[pv["highs"][-2]]["high"], data[pv["highs"][-1]]["high"]]
        rl = [data[pv["lows"][-2]]["low"], data[pv["lows"][-1]]["low"]]
        if rh[1] > rh[0] and rl[1] > rl[0]:
            tdir = "多頭"; score += 10; sigs.append(("多頭趨勢確立（高高低低）", "bull"))
        elif rh[1] < rh[0] and rl[1] < rl[0]:
            tdir = "空頭"; sigs.append(("空頭趨勢確立（低高低低）", "bear"))
        else:
            score += 2; sigs.append(("盤整區間", "neu"))
    else:
        score += 2
    if last["ma20"] is not None:
        if last["close"] > last["ma20"]:
            score += 5; sigs.append(("收盤站上20MA", "bull"))
        else:
            sigs.append(("收盤跌破20MA", "bear"))
    if last["ma60"] is not None:
        if last["close"] > last["ma60"]:
            score += 5; sigs.append(("收盤站上60MA", "bull"))
        else:
            sigs.append(("收盤跌破60MA", "bear"))
    p10 = data[-10] if len(data) >= 10 else None
    if last["ma60"] is not None and p10 and p10["ma60"] is not None:
        sl = (last["ma60"] - p10["ma60"]) / p10["ma60"] * 100
        if sl > 0.5:
            score += 5; sigs.append((f"季線向上 +{sl:.1f}%", "bull"))
        elif sl < -0.5:
            sigs.append((f"季線向下 {sl:.1f}%", "bear"))
        else:
            score += 2; sigs.append(("季線走平", "neu"))
    return {"score": min(score, 25), "max": 25, "sigs": sigs, "tdir": tdir}


def score_kline(data):
    score, sigs = 0, []
    c = data[-1]
    p1 = data[-2] if len(data) >= 2 else c
    p2 = data[-3] if len(data) >= 3 else c
    body = abs(c["close"] - c["open"])
    total = (c["high"] - c["low"]) or 0.01
    up_sh = c["high"] - max(c["close"], c["open"])
    dn_sh = min(c["close"], c["open"]) - c["low"]
    is_bull = c["close"] > c["open"]
    if is_bull:
        score += 5
        if body / total > 0.7:
            score += 3; sigs.append(("實體長紅棒", "bull"))
        else:
            sigs.append(("紅K棒", "bull"))
    else:
        if body / total > 0.7:
            sigs.append(("實體長黑棒", "bear"))
        else:
            sigs.append(("黑K棒", "bear"))
    slice20 = data[-20:]
    r_low = min(d["close"] for d in slice20)
    r_high = max(d["close"] for d in slice20)
    pos = (c["close"] - r_low) / ((r_high - r_low) or 0.01)
    if pos < 0.3:
        if dn_sh > body * 1.5:
            score += 8; sigs.append(("低檔長下影線（變盤訊號）", "bull"))
        if is_bull and c["close"] > p1["high"]:
            score += 5; sigs.append(("低檔紅K突破前高", "bull"))
    elif pos > 0.7:
        if up_sh > body * 1.5:
            sigs.append(("高檔長上影線（變盤訊號）", "bear"))
    three_bull = p2["close"] < p2["open"] and p1["close"] < p1["open"] and is_bull and c["close"] > p1["high"]
    if three_bull and pos < 0.4:
        score += 6; sigs.append(("三K底部反轉組合", "bull"))
    three_bear = p2["close"] > p2["open"] and p1["close"] > p1["open"] and (not is_bull) and c["close"] < p1["low"]
    if three_bear and pos > 0.6:
        sigs.append(("三K頂部反轉組合", "bear"))
    half = (c["high"] + c["low"]) / 2
    if is_bull and c["close"] > half:
        score += 3; sigs.append((f"收盤超過1/2價位 {half:.1f}", "bull"))
    elif (not is_bull) and c["close"] < half:
        sigs.append((f"收盤低於1/2價位 {half:.1f}", "bear"))
    return {"score": min(score, 25), "max": 25, "sigs": sigs}


def score_ma(data):
    score, sigs = 0, []
    last = data[-1]
    prev = data[-2] if len(data) >= 2 else last
    if all(last[k] is not None for k in ("ma5", "ma10", "ma20", "ma60")):
        if last["ma5"] > last["ma10"] > last["ma20"] > last["ma60"]:
            score += 10; sigs.append(("均線多頭排列", "bull"))
        elif last["ma5"] < last["ma10"] < last["ma20"] < last["ma60"]:
            sigs.append(("均線空頭排列", "bear"))
        else:
            score += 2; sigs.append(("均線糾結", "neu"))
    if len(data) >= 3:
        p1 = data[-2]
        if p1["ma5"] is not None and p1["ma5"] < p1["ma20"] and last["ma5"] > last["ma20"]:
            score += 8; sigs.append(("MA5 黃金交叉 MA20", "bull"))
        elif p1["ma5"] is not None and p1["ma5"] > p1["ma20"] and last["ma5"] < last["ma20"]:
            sigs.append(("MA5 死亡交叉 MA20", "bear"))
    if last["ma20"] is not None:
        sl20 = [d for d in data[-10:] if d["ma20"] is not None]
        slope = (sl20[-1]["ma20"] - sl20[0]["ma20"]) / sl20[0]["ma20"] * 100 if len(sl20) >= 2 else 0
        if slope > 0 and last["close"] > last["ma20"] and prev["close"] < prev["ma20"]:
            score += 5; sigs.append(("葛蘭畢買點1（突破均線）", "bull"))
        elif slope > 0 and last["close"] > last["ma20"]:
            diff = (last["close"] - last["ma20"]) / last["ma20"] * 100
            if 0 < diff < 3:
                score += 4; sigs.append(("葛蘭畢買點2（均線支撐）", "bull"))
            elif diff >= 3:
                score += 2; sigs.append(("均線上揚股價強勢", "bull"))
    if last["ma60"] is not None and last["close"] > last["ma60"]:
        score += 2; sigs.append(("股價位於季線上方", "bull"))
    return {"score": min(score, 25), "max": 25, "sigs": sigs}


def score_vol(data):
    score, sigs = 0, []
    last = data[-1]
    p1 = data[-2] if len(data) >= 2 else last
    vr = (last["volume"] / last["vm20"]) if last["vm20"] else 1
    v5r = (last["vm5"] / last["vm20"]) if (last["vm5"] and last["vm20"]) else 1
    is_up = last["close"] > p1["close"]
    if is_up:
        if vr >= 1.5:
            score += 10; sigs.append((f"上漲爆量 {vr:.1f}倍（多頭確認）", "bull"))
        elif vr >= 1.0:
            score += 6; sigs.append((f"上漲放量 {vr:.1f}倍", "bull"))
        else:
            score += 2; sigs.append(("上漲縮量（動能不足）", "neu"))
    else:
        if vr >= 1.5:
            sigs.append((f"下跌爆量 {vr:.1f}倍（賣壓沉重）", "bear"))
        elif vr >= 1.0:
            sigs.append(("下跌放量", "bear"))
        else:
            score += 5; sigs.append(("下跌縮量（賣壓減輕）", "neu"))
    low20 = min(d["low"] for d in data[-20:])
    if last["close"] < low20 * 1.15 and is_up and vr >= 1.3:
        score += 8; sigs.append(("底部放量起漲訊號", "bull"))
    if v5r > 1.2:
        score += 5; sigs.append((f"近5日均量擴增 {v5r:.1f}x", "bull"))
    elif v5r < 0.8:
        score += 1; sigs.append(("近5日均量萎縮", "neu"))
    score += 2
    return {"score": min(score, 25), "max": 25, "sigs": sigs}


# ────────────────────────────────────────────────────────────────
# DMI（趨向指標，Wilder 原始方法）＋多方力道評分（0-100，取代朱家泓四維度總分）
# ────────────────────────────────────────────────────────────────

def calc_dmi(data, period=14):
    """+DI／-DI衡量上升與下降方向的動能強弱，ADX衡量趨勢強度（不分方向），
    ADXR是ADX跟N天前ADX的平均，用來看趨勢是在增強還是減弱。用Wilder's smoothing
    對TR／+DM／-DM做平滑化（是對累計總和做平滑，不是簡單移動平均——這是DMI的
    標準算法，跟MACD用的EMA不同）。"""
    n = len(data)
    if n < period + 1:
        return None

    plus_dm, minus_dm, tr = [], [], []
    for i in range(1, n):
        up_move = data[i]["high"] - data[i - 1]["high"]
        down_move = data[i - 1]["low"] - data[i]["low"]
        plus_dm.append(up_move if (up_move > down_move and up_move > 0) else 0)
        minus_dm.append(down_move if (down_move > up_move and down_move > 0) else 0)
        tr.append(max(
            data[i]["high"] - data[i]["low"],
            abs(data[i]["high"] - data[i - 1]["close"]),
            abs(data[i]["low"] - data[i - 1]["close"]),
        ))
    if len(tr) < period:
        return None

    def wilder_sum(arr, p):
        out = []
        s = sum(arr[:p])
        out.append(s)
        for j in range(p, len(arr)):
            s = s - s / p + arr[j]
            out.append(s)
        return out

    sm_tr = wilder_sum(tr, period)
    sm_plus_dm = wilder_sum(plus_dm, period)
    sm_minus_dm = wilder_sum(minus_dm, period)

    plus_di = [(v / sm_tr[i] * 100) if sm_tr[i] else 0 for i, v in enumerate(sm_plus_dm)]
    minus_di = [(v / sm_tr[i] * 100) if sm_tr[i] else 0 for i, v in enumerate(sm_minus_dm)]
    dx = []
    for i, v in enumerate(plus_di):
        s = v + minus_di[i]
        dx.append(abs(v - minus_di[i]) / s * 100 if s else 0)

    adx = None
    if len(dx) >= period:
        adx = []
        avg = sum(dx[:period]) / period
        adx.append(avg)
        for k in range(period, len(dx)):
            avg = (avg * (period - 1) + dx[k]) / period
            adx.append(avg)

    adxr = None
    if adx and len(adx) > period:
        adxr = [(adx[m] + adx[m - period]) / 2 for m in range(period, len(adx))]

    return {
        "plusDI": plus_di[-1],
        "minusDI": minus_di[-1],
        "adx": adx[-1] if adx else None,
        "adxr": adxr[-1] if adxr else None,
    }


def score_dmi(dmi):
    """多方力道評分（0-100）：①方向性（+DI相對-DI的優勢程度，最高50分）
    ②趨勢強度（ADX，最高30分）③趨勢轉強加分（ADX>ADXR，20分）。
    DMI資料不足（新股/資料太短）時各項給0分，不是「無資料」而是保守給0分，
    因為總分要能排序／篩選，不能是非數值。"""
    if not dmi or dmi.get("plusDI") is None or dmi.get("minusDI") is None:
        return {"score": 0, "max": 100, "plusDI": None, "minusDI": None, "adx": None, "adxr": None,
                "diPts": 0, "adxPts": 0, "adxrPts": 0, "tdir": "資料不足",
                "sigs": [("DMI資料不足（可能資料天數太短）", "neu")]}

    plus_di, minus_di = dmi["plusDI"], dmi["minusDI"]
    adx, adxr = dmi.get("adx"), dmi.get("adxr")

    di_sum = plus_di + minus_di
    di_dominance_pct = (plus_di / di_sum * 100) if di_sum else 50  # 50%=中性，100%=完全多方主導
    di_pts = di_dominance_pct * 0.5  # 0-50分
    adx_pts = (min(adx, 40) / 40 * 30) if adx is not None else 0  # 0-30分，ADX≥40視為滿分
    adxr_pts = 20 if (adx is not None and adxr is not None and adx > adxr) else 0  # 0或20分
    score = round(max(0, min(100, di_pts + adx_pts + adxr_pts)))

    bullish = plus_di > minus_di
    tdir = "多頭" if bullish else ("空頭" if plus_di < minus_di else "盤整")

    sigs = [(f"+DI {plus_di:.1f}{'>' if bullish else '<'}-DI {minus_di:.1f}（多方力道{'較強' if bullish else '較弱'}）",
             "bull" if bullish else "bear")]
    if adx is not None:
        adx_lbl = "趨勢明確" if adx >= 25 else ("趨勢成形中" if adx >= 20 else "盤整")
        sigs.append((f"ADX {adx:.1f}（{adx_lbl}）", "bull" if adx >= 25 else "neu"))
    if adx is not None and adxr is not None:
        strengthening = adx > adxr
        sigs.append((f"ADX{'>' if strengthening else '<'}ADXR，趨勢{'轉強' if strengthening else '轉弱'}",
                     "bull" if strengthening else "bear"))

    return {"score": score, "max": 100, "plusDI": plus_di, "minusDI": minus_di, "adx": adx, "adxr": adxr,
            "diPts": di_pts, "adxPts": adx_pts, "adxrPts": adxr_pts, "tdir": tdir, "sigs": sigs}


def compute_core_signals(data, dm):
    """算出布林通道位置、MACD狀態、強勢突破盤／跌深反彈盤——跟 build_summary_row
    裡顯示用的邏輯完全一致，這裡抽成獨立函式回傳「原始數值」（不是格式化文字），
    給快照資料庫寫入用。刻意跟顯示邏輯分開寫（有點重複），避免改動已經在跑的
    總表顯示邏輯。"""
    last = data[-1]
    prev = data[-2] if len(data) >= 2 else last

    bb_pos = None
    if last.get("bbU") is not None and last.get("bbL") is not None:
        bb_width = last["bbU"] - last["bbL"]
        bb_pos = ((last["close"] - last["bbL"]) / bb_width * 100) if bb_width else 50

    macd_state = None
    is_breakout = False
    is_pullback_rebound = False
    if (last.get("macd") is not None and last.get("macdSig") is not None
            and last.get("macdHist") is not None and prev.get("macdHist") is not None):
        above_zero = last["macd"] > 0
        hist_growing = last["macdHist"] > prev["macdHist"]
        if above_zero and last["macdHist"] > 0:
            macd_state = "零軸上・紅柱增長" if hist_growing else "零軸上・紅柱縮短"
        elif not above_zero and last["macdHist"] < 0:
            macd_state = "零軸下・綠柱縮短" if hist_growing else "零軸下・綠柱增長"
        else:
            macd_state = "交叉轉換中"

        golden_cross_recent = False
        for gci in range(max(1, len(data) - 3), len(data)):
            gc_cur, gc_prev = data[gci], data[gci - 1]
            if (gc_cur.get("macd") is not None and gc_cur.get("macdSig") is not None
                    and gc_prev.get("macd") is not None and gc_prev.get("macdSig") is not None
                    and gc_prev["macd"] <= gc_prev["macdSig"] and gc_cur["macd"] > gc_cur["macdSig"]):
                golden_cross_recent = True
                break

        width_expanding = False
        if last.get("bbU") is not None and last.get("bbL") is not None:
            width_now_pct = (last["bbU"] - last["bbL"]) / last["close"] * 100 if last["close"] else 0
            ref_idx = len(data) - 6
            ref_bar = data[ref_idx] if ref_idx >= 0 else None
            if ref_bar and ref_bar.get("bbU") is not None and ref_bar.get("bbL") is not None and ref_bar["close"]:
                width_ref_pct = (ref_bar["bbU"] - ref_bar["bbL"]) / ref_bar["close"] * 100
                width_expanding = width_now_pct > width_ref_pct

        is_breakout = (bb_pos is not None and bb_pos >= 80 and width_expanding
                       and above_zero and last["macdHist"] > 0 and hist_growing)

        divergence_detected = False
        zz_lows = [p for p in build_zigzag(data) if p["type"] == "L"]
        if len(zz_lows) >= 2:
            recent_low, prior_low = zz_lows[-1], zz_lows[-2]
            within_lookback = recent_low["idx"] >= len(data) - 1 - 60
            macd_at_recent = data[recent_low["idx"]].get("macd") if recent_low["idx"] < len(data) else None
            macd_at_prior = data[prior_low["idx"]].get("macd") if prior_low["idx"] < len(data) else None
            if within_lookback and macd_at_recent is not None and macd_at_prior is not None:
                divergence_detected = (recent_low["price"] < prior_low["price"]) and (macd_at_recent > macd_at_prior)

        is_pullback_rebound = bb_pos is not None and bb_pos <= 20 and divergence_detected and golden_cross_recent

    return {"bb_pos": bb_pos, "macd_state": macd_state,
            "is_breakout": is_breakout, "is_pullback_rebound": is_pullback_rebound}


# ────────────────────────────────────────────────────────────────
# 快照記錄資料庫（SQLite）：每次全市場快照掃描時，把每檔股票當下的核心訊號
# 寫入一筆記錄，日積月累之後可以拿來做「預測 vs 事後實際結果」的回測分析
# （目前只做記錄，回測分析是後續的事）。
# ────────────────────────────────────────────────────────────────
SNAPSHOT_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "snapshots.db")


def init_snapshot_db():
    conn = sqlite3.connect(SNAPSHOT_DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS snapshots (
            snap_date TEXT NOT NULL,
            stock_id TEXT NOT NULL,
            name TEXT,
            score INTEGER,
            plus_di REAL,
            minus_di REAL,
            adx REAL,
            adxr REAL,
            bb_pos REAL,
            macd_state TEXT,
            is_breakout INTEGER,
            is_pullback_rebound INTEGER,
            close REAL,
            volume REAL,
            chg_pct REAL,
            created_at TEXT,
            PRIMARY KEY (snap_date, stock_id)
        )
    """)
    conn.commit()
    return conn


def save_snapshot_row(conn, snap_date, stock_id, name, dm, sig, close, volume, chg_pct):
    conn.execute("""
        INSERT OR REPLACE INTO snapshots
        (snap_date, stock_id, name, score, plus_di, minus_di, adx, adxr, bb_pos, macd_state,
         is_breakout, is_pullback_rebound, close, volume, chg_pct, created_at)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (
        snap_date, stock_id, name, dm["score"], dm["plusDI"], dm["minusDI"], dm["adx"], dm["adxr"],
        sig["bb_pos"], sig["macd_state"], int(sig["is_breakout"]), int(sig["is_pullback_rebound"]),
        close, volume, chg_pct, datetime.now().isoformat(timespec="seconds"),
    ))


def get_snapshot_stats():
    if not os.path.exists(SNAPSHOT_DB_PATH):
        return None
    conn = sqlite3.connect(SNAPSHOT_DB_PATH)
    try:
        row = conn.execute(
            "SELECT COUNT(*), COUNT(DISTINCT snap_date), COUNT(DISTINCT stock_id), MAX(snap_date), MIN(snap_date) FROM snapshots"
        ).fetchone()
        return {"total_rows": row[0], "distinct_dates": row[1], "distinct_stocks": row[2],
                "latest_date": row[3], "earliest_date": row[4]}
    except sqlite3.OperationalError:
        # 資料庫檔案存在，但 snapshots 表還沒建立（還沒跑過第一次快照掃描）
        return None
    finally:
        conn.close()


def run_full_market_snapshot(token: str):
    """全市場快照掃描：固定用『上市清單』（TWSE_LIST，約1100+檔）當樣本池，
    不是目前輸入框裡的清單——樣本數愈大，之後回測的統計意義才夠。只算核心的
    DMI/布林通道/MACD訊號（不勾P/E、營收、法人這些額外欄位），避免全市場規模
    下API耗用跟耗時暴增。這是長時間操作（1100多檔，預估10-20分鐘），過程中
    請不要切換頁面或關閉分頁。"""
    conn = init_snapshot_db()
    snap_date = datetime.today().strftime("%Y-%m-%d")
    name_map = fetch_stock_name_map(token)

    total = len(TWSE_LIST)
    progress_bar = st.progress(0)
    status = st.empty()
    log_box = st.expander("📸 快照掃描紀錄（展開查看逐檔進度）", expanded=False)
    saved, failed = 0, 0

    for i, sid in enumerate(TWSE_LIST):
        status.text(f"📸 快照中：{sid}… ({i + 1}/{total})　已存 {saved}　失敗 {failed}")
        try:
            rows = fetch_price_data(sid, token, 180)  # 180天足夠算DMI(14)+背離(60天回看)
            raw_data = sorted(
                [{"date": d["date"], "open": float(d["open"]), "high": float(d["max"]),
                  "low": float(d["min"]), "close": float(d["close"]), "volume": float(d["Trading_Volume"])}
                 for d in rows],
                key=lambda x: x["date"],
            )
            data = enrich(raw_data)
            dmi = calc_dmi(data, 14)
            dm = score_dmi(dmi)
            sig = compute_core_signals(data, dm)
            last = data[-1]
            prev = data[-2] if len(data) >= 2 else last
            chgp = (last["close"] - prev["close"]) / prev["close"] * 100 if prev["close"] else 0
            name = name_map.get(sid, sid)
            save_snapshot_row(conn, snap_date, sid, name, dm, sig, last["close"], last["volume"], chgp)
            saved += 1
            if saved % 50 == 0:
                conn.commit()  # 每50檔提交一次，中途萬一中斷也不會全部流失
        except Exception as ex:
            failed += 1
            with log_box:
                st.caption(f"❌ {sid} 失敗：{ex}")
        progress_bar.progress((i + 1) / total)
        if i < total - 1:
            time.sleep(0.3)

    conn.commit()
    conn.close()
    progress_bar.empty()
    status.empty()
    st.success(f"✅ {snap_date} 全市場快照完成：成功 {saved} 檔，失敗 {failed} 檔（共 {total} 檔）")


# ────────────────────────────────────────────────────────────────
# 歷史回測（回溯過去N個月＋事後N天報酬驗證＋參數網格搜尋）
# 用同一個資料庫檔案（SNAPSHOT_DB_PATH），存到另一張表 backtest_evals。
# ────────────────────────────────────────────────────────────────
BACKTEST_HORIZONS = [5, 10, 20]  # 事後驗證用的天數（皆為交易日）


def init_backtest_table():
    conn = sqlite3.connect(SNAPSHOT_DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS backtest_evals (
            eval_date TEXT NOT NULL,
            stock_id TEXT NOT NULL,
            name TEXT,
            score INTEGER,
            plus_di REAL,
            minus_di REAL,
            adx REAL,
            adxr REAL,
            bb_pos REAL,
            is_breakout INTEGER,
            is_pullback_rebound INTEGER,
            entry_close REAL,
            ret_5d REAL,
            ret_10d REAL,
            ret_20d REAL,
            created_at TEXT,
            PRIMARY KEY (eval_date, stock_id)
        )
    """)
    conn.commit()
    return conn


def run_historical_backtest(token: str, months_back: int = 3):
    """回溯過去 months_back 個月，對每個交易日重新計算「當時」的DMI/評分/訊號
    （只用當天以前的資料切片再丟進既有的 calc_dmi／score_dmi／compute_core_signals，
    保證沒有偷看未來——這幾個函式本來就是「給一段資料、算出最後一天的訊號」，
    直接重複利用，不用另外寫一套向量化版本冒index算錯的風險），然後對照
    5/10/20天後的實際收盤價算出報酬率，存進 backtest_evals 表。

    每檔股票抓約(3個月+60天技術指標暖身緩衝+20天事後驗證緩衝)的歷史，固定用
    上市清單當樣本池。運算量本身很小（純迴圈，沒有額外API呼叫），真正花時間的
    還是1,100多檔的價格資料抓取。
    """
    conn = init_backtest_table()
    name_map = fetch_stock_name_map(token)
    lookback_buffer = 60
    max_horizon = max(BACKTEST_HORIZONS)
    fetch_days = months_back * 30 + lookback_buffer + max_horizon + 10

    total = len(TWSE_LIST)
    progress_bar = st.progress(0)
    status = st.empty()
    log_box = st.expander("🔬 回測進度紀錄（展開查看逐檔進度）", expanded=False)
    saved_rows, failed = 0, 0

    for i, sid in enumerate(TWSE_LIST):
        status.text(f"🔬 回測中：{sid}… ({i + 1}/{total})　已存 {saved_rows:,} 筆　失敗 {failed} 檔")
        try:
            rows = fetch_price_data(sid, token, fetch_days)
            raw_data = sorted(
                [{"date": d["date"], "open": float(d["open"]), "high": float(d["max"]),
                  "low": float(d["min"]), "close": float(d["close"]), "volume": float(d["Trading_Volume"])}
                 for d in rows],
                key=lambda x: x["date"],
            )
            data = enrich(raw_data)
            n = len(data)
            eval_start = lookback_buffer
            eval_end = n - max_horizon  # 不含此index，確保每個評估點後面都還有滿20天可以驗證
            if eval_end <= eval_start:
                failed += 1
                continue
            name = name_map.get(sid, sid)
            for idx in range(eval_start, eval_end):
                data_slice = data[:idx + 1]  # 只給「當時」以前的資料，不含未來
                dmi = calc_dmi(data_slice, 14)
                dm = score_dmi(dmi)
                sig = compute_core_signals(data_slice, dm)
                entry_close = data[idx]["close"]
                rets = {}
                for h in BACKTEST_HORIZONS:
                    if idx + h < n and entry_close:
                        rets[h] = (data[idx + h]["close"] - entry_close) / entry_close * 100
                    else:
                        rets[h] = None
                conn.execute("""
                    INSERT OR REPLACE INTO backtest_evals
                    (eval_date, stock_id, name, score, plus_di, minus_di, adx, adxr, bb_pos,
                     is_breakout, is_pullback_rebound, entry_close, ret_5d, ret_10d, ret_20d, created_at)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    data[idx]["date"], sid, name, dm["score"], dm["plusDI"], dm["minusDI"], dm["adx"], dm["adxr"],
                    sig["bb_pos"], int(sig["is_breakout"]), int(sig["is_pullback_rebound"]), entry_close,
                    rets[5], rets[10], rets[20], datetime.now().isoformat(timespec="seconds"),
                ))
                saved_rows += 1
            if (i + 1) % 20 == 0:
                conn.commit()
        except Exception as ex:
            failed += 1
            with log_box:
                st.caption(f"❌ {sid} 失敗：{ex}")
        progress_bar.progress((i + 1) / total)
        if i < total - 1:
            time.sleep(0.3)

    conn.commit()
    conn.close()
    progress_bar.empty()
    status.empty()
    st.success(f"✅ 歷史回測完成：{saved_rows:,} 筆評估紀錄（{total - failed}/{total} 檔成功），可以往下看分析結果")


def load_backtest_df():
    if not os.path.exists(SNAPSHOT_DB_PATH):
        return None
    conn = sqlite3.connect(SNAPSHOT_DB_PATH)
    try:
        df = pd.read_sql_query("SELECT * FROM backtest_evals", conn)
    except Exception:
        df = None
    conn.close()
    if df is None or df.empty:
        return None
    return df


def analyze_score_buckets(df: pd.DataFrame) -> pd.DataFrame:
    def bucket(s):
        if s >= 80:
            return "80-100（積極做多）"
        if s >= 65:
            return "65-79（可考慮進場）"
        if s >= 50:
            return "50-64（觀望）"
        return "0-49（不建議）"

    d = df.copy()
    d["bucket"] = d["score"].apply(bucket)
    order = ["0-49（不建議）", "50-64（觀望）", "65-79（可考慮進場）", "80-100（積極做多）"]
    rows = []
    for b, g in d.groupby("bucket"):
        row = {"評分區間": b, "樣本數": len(g)}
        for h in BACKTEST_HORIZONS:
            valid = g[f"ret_{h}d"].dropna()
            row[f"{h}日平均報酬%"] = round(valid.mean(), 2) if len(valid) else None
            row[f"{h}日勝率%"] = round((valid > 0).mean() * 100, 1) if len(valid) else None
        rows.append(row)
    result = pd.DataFrame(rows)
    result["_order"] = result["評分區間"].apply(lambda x: order.index(x) if x in order else 99)
    return result.sort_values("_order").drop(columns="_order").reset_index(drop=True)


def analyze_tag_hitrate(df: pd.DataFrame, tag_col: str, tag_label: str) -> pd.DataFrame:
    rows = []
    for val, g in df.groupby(tag_col):
        label = f"{tag_label}＝是" if val == 1 else f"{tag_label}＝否"
        row = {"標記": label, "樣本數": len(g)}
        for h in BACKTEST_HORIZONS:
            valid = g[f"ret_{h}d"].dropna()
            row[f"{h}日平均報酬%"] = round(valid.mean(), 2) if len(valid) else None
            row[f"{h}日勝率%"] = round((valid > 0).mean() * 100, 1) if len(valid) else None
        rows.append(row)
    return pd.DataFrame(rows)


def grid_search_params(df: pd.DataFrame, target_horizon: int = 10):
    """在已收集的原始指標值（+DI/-DI/ADX/ADXR）上，重新代入不同的評分公式參數
    （方向性權重、ADX滿分上限、ADXR加分值）試算，看哪組參數對N天後報酬的判斷力
    比較好——不需要重新抓資料，純粹是在同一份歷史資料上換公式重算。
    這是簡化版網格搜尋，樣本數有限時容易過度適配歷史資料，結果僅供參考方向，
    不建議未經檢視就直接套用到正式評分公式。"""
    ret_col = f"ret_{target_horizon}d"
    valid_df = df.dropna(subset=[ret_col, "plus_di", "minus_di", "adx"]).copy()
    if valid_df.empty:
        return pd.DataFrame()

    di_sum = valid_df["plus_di"] + valid_df["minus_di"]
    di_dom = np.where(di_sum > 0, valid_df["plus_di"] / di_sum * 100, 50)
    has_adxr = valid_df["adxr"].notna()
    adx_gt_adxr = valid_df["adx"] > valid_df["adxr"].fillna(-1)

    results = []
    for adx_cap in [30, 40, 50]:
        for adxr_bonus in [15, 20, 25]:
            for di_weight in [0.4, 0.5, 0.6]:
                di_pts = di_dom * di_weight
                adx_max_pts = max(0, 100 - 100 * di_weight - adxr_bonus)  # 剩餘配分給ADX，確保三項頂多加到100
                adx_pts = np.minimum(valid_df["adx"], adx_cap) / adx_cap * adx_max_pts
                adxr_pts = np.where(has_adxr & adx_gt_adxr, adxr_bonus, 0)
                new_score = np.clip(di_pts + adx_pts + adxr_pts, 0, 100)

                for buy_thr in [50, 65, 80]:
                    buy_mask = new_score >= buy_thr
                    if buy_mask.sum() < 20:
                        continue
                    rets = valid_df.loc[buy_mask, ret_col]
                    results.append({
                        "方向性權重": di_weight, "ADX滿分上限": adx_cap, "ADXR加分": adxr_bonus,
                        "買進門檻": buy_thr, "訊號數": int(buy_mask.sum()),
                        f"{target_horizon}日平均報酬%": round(rets.mean(), 2),
                        f"{target_horizon}日勝率%": round((rets > 0).mean() * 100, 1),
                    })
    result_df = pd.DataFrame(results)
    if result_df.empty:
        return result_df
    return result_df.sort_values(f"{target_horizon}日平均報酬%", ascending=False).head(15).reset_index(drop=True)


# ────────────────────────────────────────────────────────────────
# 回後買上漲 8 條件核對
# ────────────────────────────────────────────────────────────────

def check_pullback_buy(data):
    last = data[-1]
    prev = data[-2] if len(data) >= 2 else last

    results = []
    all_pass = True

    # 轉折波（跟型態辨識、圖表轉折波用同一套 build_zigzag，不再用另一套獨立的分形判斷法）
    zz = build_zigzag(data)
    zz_highs = [p for p in zz if p["type"] == "H"]
    zz_lows = [p for p in zz if p["type"] == "L"]

    # ── 條件1：趨勢多頭（轉折波 高高低低）──
    c1 = False
    if len(zz_highs) >= 2 and len(zz_lows) >= 2:
        rh = [zz_highs[-2]["price"], zz_highs[-1]["price"]]
        rl = [zz_lows[-2]["price"], zz_lows[-1]["price"]]
        c1 = rh[1] > rh[0] and rl[1] > rl[0]
    results.append({"label": "①趨勢多頭（高高低低）", "pass": c1, "required": True, "detail": ""})
    if not c1:
        all_pass = False

    # ── 條件2：位置回後上漲（前1~5日曾出現縮量回檔，今日反彈）──
    # 依課程講義：回檔幅度用費波那契回檔比例分級，回檔愈淺（愈接近0.382）代表股票愈強，
    # 愈要把握機會進場；回檔愈深（接近或超過0.618）力道愈弱。
    pullback_days = data[-6:-1]
    had_pullback = any(
        (d["close"] < (pullback_days[i - 1]["close"] if i > 0 else d["close"])) or (d["close"] < d["open"])
        for i, d in enumerate(pullback_days)
    )
    c2 = had_pullback and last["close"] > prev["close"]
    pullback_detail = ""
    if zz_highs:
        peak_point = zz_highs[-1]
        prior_low_cands = [p for p in zz_lows if p["idx"] < peak_point["idx"]]
        if prior_low_cands:
            prior_low = prior_low_cands[-1]["price"]
            peak = peak_point["price"]
            if peak > prior_low:
                pullback_segment = data[peak_point["idx"]:]
                pullback_low = min(d["low"] for d in pullback_segment)
                retrace = (peak - pullback_low) / (peak - prior_low)
                grade = "最強" if retrace <= 0.382 else "強" if retrace <= 0.5 else "弱" if retrace <= 0.618 else "回檔過深"
                pullback_detail = f"回檔幅度{retrace * 100:.1f}%（{grade}，費波0.382/0.5/0.618分級）"
    results.append({"label": "②位置回後上漲（近期有回檔，今轉上）", "pass": c2, "required": True, "detail": pullback_detail})
    if not c2:
        all_pass = False

    c3 = last["ma5"] is not None and last["close"] > last["ma5"]
    results.append({"label": "③收盤站上5MA（平價不算）", "pass": c3, "required": True,
                     "detail": f"5MA={last['ma5']:.2f}  收盤={last['close']:.2f}" if last["ma5"] is not None else ""})
    if not c3:
        all_pass = False

    c4 = last["high"] > prev["high"]
    results.append({"label": "④突破前一日高點（含上影線）", "pass": c4, "required": True,
                     "detail": f"今高={last['high']:.2f}  昨高={prev['high']:.2f}"})
    if not c4:
        all_pass = False

    chg_pct = (last["close"] - prev["close"]) / prev["close"] * 100 if prev["close"] > 0 else 0
    c5 = chg_pct >= 2.0
    results.append({"label": "⑤漲幅2%以上", "pass": c5, "required": True, "detail": f"漲幅={chg_pct:.2f}%"})
    if not c5:
        all_pass = False

    body = last["close"] - last["open"]
    up_sh = last["high"] - last["close"]
    dn_sh = last["open"] - last["low"]
    max_sh = max(up_sh, dn_sh)
    c6 = body > 0 and max_sh <= body
    results.append({"label": "⑥實體紅K，影線不大於實體", "pass": c6, "required": True,
                     "detail": f"實體={body:.2f}  最大影線={max_sh:.2f}"})
    if not c6:
        all_pass = False

    vol_ratio = (last["volume"] / last["vm20"]) if last["vm20"] else 1
    c7 = vol_ratio >= 1.0
    results.append({"label": "⑦成交量增（加分項）", "pass": c7, "required": False, "detail": f"量比MA20={vol_ratio:.2f}x"})

    prev_kd = data[-2] if len(data) >= 2 else None
    c8 = False
    if last["kdK"] is not None and prev_kd is not None and prev_kd["kdK"] is not None:
        c8 = last["kdK"] > prev_kd["kdK"]
    kd_detail = (f"K={last['kdK']:.1f}  昨K={prev_kd['kdK']:.1f}" if (last["kdK"] is not None and prev_kd and prev_kd["kdK"] is not None) else "KD資料不足")
    results.append({"label": "⑧指標確認（K值向上）", "pass": c8, "required": True, "detail": kd_detail})
    if not c8:
        all_pass = False

    required_total = sum(1 for r in results if r["required"])
    required_passed = sum(1 for r in results if r["required"] and r["pass"])
    bonus_passed = sum(1 for r in results if not r["required"] and r["pass"])

    return {"results": results, "allPass": all_pass, "requiredPassed": required_passed,
            "requiredTotal": required_total, "bonusPassed": bonus_passed}


# ────────────────────────────────────────────────────────────────
# 型態確認：朱家泓 進場型態（6種底部型態＋ABC切線＋上升軌道＋大量黑K＋回後買上漲）
# (1)頭肩底 (2)複式頭肩底 (3)N字底 (4)三重底 (5)圓弧底 (6)一字底(均線糾結)
# (7)突破ABC修正下降切線 (8)突破上升軌道線 (9)突破飆股大量黑K最高點 (10)回後買上漲
# ────────────────────────────────────────────────────────────────

def tolerant(a, b, pct):
    base = max(abs(a), abs(b), 1e-6)
    return abs(a - b) / base <= pct


def _ma_slope_up(data, key, n, last_idx):
    p_idx = max(0, last_idx - n)
    p, c = data[p_idx], data[last_idx]
    if c[key] is None or p[key] is None:
        return False
    return c[key] > p[key]


def detect_patterns(data, pb, skip_just_broke=False):
    # 型態辨識固定看跟圖表一致的完整資料範圍（也就是「分析天數」實際抓到的全部K棒），
    # 並套用同一套壞資料過濾規則，確保型態辨識用的轉折波，跟圖表上實際畫出來的
    # 轉折波／輔助線是同一組資料算出來的結果——避免用不同範圍算出對不起來的轉折點。
    data = [d for d in data if d.get("open", 0) > 0 and d.get("high", 0) > 0
            and d.get("low", 0) > 0 and d.get("close", 0) > 0
            and all(_is_finite(d[k]) for k in ("open", "high", "low", "close"))]
    last = len(data) - 1
    last_close = data[last]["close"]
    last_vol = data[last]["volume"]
    vm20 = data[last]["vm20"]
    vol_confirm = (last_vol / vm20 >= 1.3) if vm20 else False

    # 型態辨識所用的高低點，改成直接沿用「轉折波」(build_zigzag) 算出來的同一組轉折點，
    # 不再用另一套獨立的分形視窗判斷法——這樣圖表上畫出來的轉折波，就是型態辨識實際依據的高低點，
    # 兩者完全一致，不會有「圖上看到的轉折」跟「型態判斷用的轉折」對不起來的狀況。
    zz = build_zigzag(data)
    # 型態辨識用的轉折點範圍，直接沿用整個(已限制在120根K棒內的)資料範圍，跟圖表顯示範圍完全一致，
    # 不再另外疊加一層90天子視窗限制——避免「圖表上看得到的高低點」卻被排除在型態判斷之外，
    # 導致畫出來的頸線/壓力線跟圖上真正的高低點對不起來。
    floor = 0
    lows = [p["idx"] for p in zz if p["type"] == "L" and floor <= p["idx"] < last]
    highs = [p["idx"] for p in zz if p["type"] == "H" and floor <= p["idx"] < last]

    def breakout_check(resistance):
        if resistance is None:
            return {"confirmed": False, "detail": ""}
        return {
            "confirmed": last_close > resistance,
            "detail": f"頸線/壓力＝{resistance:.2f}　現價＝{last_close:.2f}" + ("　(帶量突破)" if vol_confirm else ""),
        }

    results = []

    # (1) 頭肩底：右肩不破 頭→頸線 1/2（依課程講義）
    id_, name = "hs", "頭肩底"
    added = False
    if len(lows) >= 3:
        l3 = lows[-3:]
        L1, L2, L3v = data[l3[0]]["low"], data[l3[1]]["low"], data[l3[2]]["low"]
        shoulders_similar = tolerant(L1, L3v, 0.06)
        head_lower = L2 < L1 * 0.985 and L2 < L3v * 0.985
        if shoulders_similar and head_lower:
            h_between = [i for i in highs if l3[0] < i < l3[2]]
            neck = max((data[i]["high"] for i in h_between), default=None)
            # 依課程講義：右肩(L3)不能跌破 頭(L2)→頸線 漲幅的 1/2，才是有效的右肩（不是單純比頭高就好）
            valid_right_shoulder = neck is not None and L3v > (L2 + neck) / 2
            if valid_right_shoulder:
                bo = breakout_check(neck)
                results.append({"id": id_, "name": name, "formed": True, "breakout": bo["confirmed"],
                                 "detail": bo["detail"] or "型態成形，等待突破頸線", "desc": "左右肩低點相近，頭部最低，右肩不破1/2，突破頸線為買點",
                                 "line": {"i1": l3[0], "p1": neck, "slope": 0}})
                added = True
    if not added:
        results.append({"id": id_, "name": name, "formed": False, "breakout": False,
                         "detail": "尚未偵測到符合結構", "desc": "左右肩低點相近，頭部最低，右肩不破1/2，突破頸線為買點"})

    # (2) 複式頭肩底：最近肩部不破 頭→頸線 1/2（依課程講義）
    id_, name = "chs", "複式頭肩底"
    added = False
    if len(lows) >= 4:
        last_lows = lows[-5:]
        low_vals = [data[i]["low"] for i in last_lows]
        min_val = min(low_vals)
        head_pos = low_vals.index(min_val)
        has_left = head_pos > 0
        has_right = head_pos < len(last_lows) - 1
        shoulder_vals = [v for idx, v in enumerate(low_vals) if idx != head_pos]
        shoulders_ok = has_left and has_right and shoulder_vals and all(
            tolerant(v, shoulder_vals[0], 0.08) and v > min_val * 1.02 for v in shoulder_vals
        )
        if shoulders_ok:
            h_between = [i for i in highs if last_lows[0] < i < last_lows[-1]]
            neck = max((data[i]["high"] for i in h_between), default=None)
            # 依課程講義：最近（最右）一個肩部不能跌破 頭→頸線 漲幅的 1/2
            last_shoulder_val = low_vals[-1]
            valid_last_shoulder = neck is not None and last_shoulder_val > (min_val + neck) / 2
            if valid_last_shoulder:
                bo = breakout_check(neck)
                results.append({"id": id_, "name": name, "formed": True, "breakout": bo["confirmed"],
                                 "detail": bo["detail"] or "型態成形，等待突破頸線", "desc": "多重肩部低點環繞單一最低頭部，最近肩部不破1/2，突破頸線為買點",
                                 "line": {"i1": last_lows[0], "p1": neck, "slope": 0}})
                added = True
    if not added:
        results.append({"id": id_, "name": name, "formed": False, "breakout": False,
                         "detail": "尚未偵測到符合結構", "desc": "多重肩部低點環繞單一最低頭部，最近肩部不破1/2，突破頸線為買點"})

    # (3) N字底：依課程講義，拉回不破 A→B 漲幅的 1/2 才是有效的淺拉回
    id_, name = "nb", "N字底"
    added = False
    if len(lows) >= 2 and len(highs) >= 1:
        A, C = lows[-2], lows[-1]
        b_cands = [i for i in highs if A < i < C]
        if b_cands:
            # 取A、C之間「最高」的確認高點，而不是「最近」的一個——
            # 若下跌過程中出現次要反彈小高點，會比真正的主高點更晚被確認，
            # 用「最近」選到的話會抓到錯誤（偏低）的壓力位置。
            B = max(b_cands, key=lambda i: data[i]["high"])
            low_a, low_c, high_b = data[A]["low"], data[C]["low"], data[B]["high"]
            # 依課程講義：C（拉回低點）不能跌破 A→B 漲幅的 1/2，才是有效的淺拉回（不是單純比A高就好）
            half_point = (low_a + high_b) / 2
            shallow_pullback = low_c > half_point
            if shallow_pullback:
                bo = breakout_check(high_b)
                results.append({"id": id_, "name": name, "formed": True, "breakout": bo["confirmed"],
                                 "detail": bo["detail"] or f"拉回未破1/2（{half_point:.2f}），等待突破反彈高點", "desc": "低點反彈後拉回不破1/2，再突破反彈高點為買點",
                                 "line": {"i1": B, "p1": high_b, "slope": 0}})
                added = True
    if not added:
        results.append({"id": id_, "name": name, "formed": False, "breakout": False,
                         "detail": "尚未偵測到符合結構", "desc": "低點反彈後拉回不破1/2，再突破反彈高點為買點"})

    # (4) 三重底：三個相近低點，兩個中間高點也要相近（形成真正的水平頸線）
    id_, name = "tb", "三重底"
    added = False
    if len(lows) >= 3:
        l3 = lows[-3:]
        vals = [data[i]["low"] for i in l3]
        all_similar = tolerant(vals[0], vals[1], 0.08) and tolerant(vals[1], vals[2], 0.08) and tolerant(vals[0], vals[2], 0.08)
        if all_similar:
            h_between1 = [i for i in highs if l3[0] < i < l3[1]]
            h_between2 = [i for i in highs if l3[1] < i < l3[2]]
            peak1 = max((data[i]["high"] for i in h_between1), default=None)
            peak2 = max((data[i]["high"] for i in h_between2), default=None)
            # 依課程講義：兩個中間高點要相近，才是真正的水平頸線（不是隨便夾兩個高低不一的高點）
            neckline_ok = peak1 is not None and peak2 is not None and tolerant(peak1, peak2, 0.05)
            if neckline_ok:
                res = max(peak1, peak2)
                bo = breakout_check(res)
                results.append({"id": id_, "name": name, "formed": True, "breakout": bo["confirmed"],
                                 "detail": bo["detail"] or "型態成形，等待突破壓力", "desc": "三個低點高度相近，中間兩高點形成水平頸線，突破頸線為買點",
                                 "line": {"i1": l3[0], "p1": res, "slope": 0}})
                added = True
    if not added:
        results.append({"id": id_, "name": name, "formed": False, "breakout": False,
                         "detail": "尚未偵測到符合結構", "desc": "三個低點高度相近，中間兩高點形成水平頸線，突破頸線為買點"})

    # (5) 圓弧底
    id_, name = "rb", "圓弧底"
    added = False
    win = data[-40:]
    if len(win) >= 30:
        seg = len(win) // 3
        first, mid, tail = win[:seg], win[seg:len(win) - seg], win[len(win) - seg:]

        def avg(arr, key):
            return sum(d[key] for d in arr) / len(arr)

        def slope(arr):
            n = len(arr)
            sx = sy = sxy = sxx = 0
            for i, d in enumerate(arr):
                sx += i; sy += d["close"]; sxy += i * d["close"]; sxx += i * i
            denom = (n * sxx - sx * sx) or 1
            return (n * sxy - sx * sy) / denom

        slope_first, slope_tail = slope(first), slope(tail)
        mid_low = min(d["low"] for d in mid)
        is_convex = avg(first, "close") > mid_low * 1.01 and avg(tail, "close") > mid_low * 1.01
        shape_ok = slope_first < 0 and slope_tail > 0 and is_convex
        avg_range = sum((d["high"] - d["low"]) / d["close"] for d in win) / len(win)
        low_vol = avg_range < 0.05
        if shape_ok and low_vol:
            resistance = max(d["high"] for d in first)
            bo = breakout_check(resistance)
            win_start_idx = len(data) - len(win)
            # 依課程講義：目標價 = 突破點 + 型態高度（起跌點高點 - 最低點），即測量移動法
            target = resistance + (resistance - mid_low)
            results.append({"id": id_, "name": name, "formed": True, "breakout": bo["confirmed"],
                             "detail": (bo["detail"] or "弧形築底中，等待突破起跌壓力") + f"　目標價≈{target:.2f}",
                             "desc": "價格緩跌後緩升成U型，突破起跌點高點為買點",
                             "line": {"i1": win_start_idx, "p1": resistance, "slope": 0}})
            added = True
    if not added:
        results.append({"id": id_, "name": name, "formed": False, "breakout": False,
                         "detail": "尚未偵測到符合結構", "desc": "價格緩跌後緩升成U型，突破起跌點高點為買點"})

    # (6) 一字底（均線糾結）：整理區間範圍約10%內（依課程講義定義）
    id_, name = "fb", "一字底(均線糾結)"
    added = False
    N = 10
    win = data[-N - 1:-1]
    ok = len(win) == N and all(d["ma5"] is not None and d["ma10"] is not None and d["ma20"] is not None and d["ma60"] is not None for d in win)
    if ok:
        tangled = all(
            (max(d["ma5"], d["ma10"], d["ma20"], d["ma60"]) - min(d["ma5"], d["ma10"], d["ma20"], d["ma60"]))
            / min(d["ma5"], d["ma10"], d["ma20"], d["ma60"]) <= 0.05
            for d in win
        )
        # 課程講義定義：整理區間範圍（整段最高與最低價的差距）約在 10% 以內，才算「一字底」
        win_high = max(d["high"] for d in win)
        win_low = min(d["low"] for d in win)
        narrow_range = (win_high - win_low) / win_low <= 0.10
        if tangled and narrow_range:
            resistance = win_high
            last4 = [data[last]["ma5"], data[last]["ma10"], data[last]["ma20"], data[last]["ma60"]]
            above_all_ma = all(v is not None and last_close > v for v in last4)
            breakout = last_close > resistance and above_all_ma and vol_confirm
            win_start_idx6 = len(data) - N - 1
            # 依課程講義：目標價 = 突破點 + 型態高度（整理區間高點 - 整理區間低點）
            target6 = resistance + (resistance - win_low)
            results.append({"id": id_, "name": name, "formed": True, "breakout": breakout,
                             "detail": f"整理區間高點＝{resistance:.2f}　現價＝{last_close:.2f}" + ("　(帶量突破)" if vol_confirm else "　(尚未帶量)") + f"　目標價≈{target6:.2f}",
                             "desc": "均線糾結、價格窄幅整理（區間範圍約10%內），帶量突破整理區間為買點",
                             "line": {"i1": win_start_idx6, "p1": resistance, "slope": 0}})
            added = True
    if not added:
        results.append({"id": id_, "name": name, "formed": False, "breakout": False,
                         "detail": "尚未偵測到符合結構", "desc": "均線糾結、價格窄幅整理（區間範圍約10%內），帶量突破整理區間為買點"})

    # (7) 突破ABC修正下降切線：A、B兩高點畫下降切線，B之後要有C段拉回低點
    id_, name = "abc", "突破ABC修正下降切線"
    added = False
    # 只看「最近」的轉折高點，取最後兩個（A、B），避免抓到太久遠、已經沒有參考意義的舊高點
    recent_highs = [i for i in highs if i >= last - 40]
    if len(recent_highs) >= 2:
        h1, h2 = recent_highs[-2], recent_highs[-1]  # A：較早較高；B：較近較低
        y1, y2 = data[h1]["high"], data[h2]["high"]
        # A、B之間要有拉回的低點（確認A→低點→B是有效的一次反彈，不是隨便兩個高點連線）
        has_low_between = any(h1 < li < h2 for li in lows)
        # 依課程講義：B之後還要有一段「C」低點（真正的拉回），突破才是站在C低點之上完成的，
        # 不能B之後價格根本沒拉回就直接算突破——那樣就只是A、B兩個高點連線，不是完整的ABC三段修正。
        c_low = None
        for ci in range(h2 + 1, last):
            if c_low is None or data[ci]["low"] < c_low:
                c_low = data[ci]["low"]
        # 至少要有2%以上的拉回幅度，才算真正的C段（避免單純的價格雜訊被誤判成有效拉回）
        has_c_leg = c_low is not None and c_low < data[h2]["close"] * 0.98
        if y2 < y1 and h2 > h1 and has_low_between and has_c_leg and (h2 - h1) <= 20 and (last - h2) <= 20:
            slope_ = (y2 - y1) / (h2 - h1)
            line_at_last = y1 + slope_ * (last - h1)
            ma20up = _ma_slope_up(data, "ma20", 10, last)
            is_red = data[last]["close"] > data[last]["open"]
            breakout = last_close > line_at_last and ma20up and is_red
            results.append({"id": id_, "name": name, "formed": True, "breakout": breakout,
                             "detail": f"下降切線位置≈{line_at_last:.2f}　C低點＝{c_low:.2f}　現價＝{last_close:.2f}" + ("　MA20上揚" if ma20up else "　MA20未上揚"),
                             "desc": "多頭回檔呈ABC三段式下跌，A、B高點畫下降切線，C段拉回後帶量紅K突破切線為買點",
                             "line": {"i1": h1, "p1": y1, "i2": h2, "p2": y2, "slope": slope_}})
            added = True
    if not added:
        results.append({"id": id_, "name": name, "formed": False, "breakout": False,
                         "detail": "尚未偵測到符合結構", "desc": "多頭回檔呈ABC三段式下跌，A、B高點畫下降切線，C段拉回後帶量紅K突破切線為買點"})

    # (8) 突破上升軌道線：軌道線要有至少2個高點貼著同一條平行線才算數（依課程講義）
    id_, name = "channel", "突破上升軌道線"
    added = False
    floor2 = max(0, len(data) - 60)
    lows_in = [i for i in lows if i >= floor2]
    highs_in = [i for i in highs if i >= floor2]
    if len(lows_in) >= 2:
        l1, l2 = lows_in[-2], lows_in[-1]
        ly1, ly2 = data[l1]["low"], data[l2]["low"]
        if ly2 > ly1 and l2 > l1:
            slope2 = (ly2 - ly1) / (l2 - l1)
            offset_cands = [data[i]["high"] - (ly1 + slope2 * (i - l1)) for i in highs_in if i > l1]
            # 依課程講義：上升軌道線要有「至少2個高點」貼著同一條平行線，才是真正的軌道線
            # （不能只是單一個高點恰好離支撐線最遠，那只是巧合，不是真的通道）
            # 容忍度用股價的3%來算（而不是用offset自身的百分比），避免offset數值太小時誤判過嚴
            offset = None
            if len(offset_cands) >= 2:
                tol = last_close * 0.03
                best_offset, best_count = None, 0
                for o in offset_cands:
                    count = sum(1 for o2 in offset_cands if abs(o - o2) <= tol)
                    if count > best_count or (count == best_count and (best_offset is None or o > best_offset)):
                        best_count, best_offset = count, o
                if best_count >= 2:
                    offset = best_offset
            if offset is not None and offset > 0:
                upper_at_last = ly1 + slope2 * (last - l1) + offset
                ma20up2 = _ma_slope_up(data, "ma20", 10, last)
                is_red2 = data[last]["close"] > data[last]["open"]
                breakout2 = last_close > upper_at_last and ma20up2 and is_red2 and vol_confirm
                results.append({"id": id_, "name": name, "formed": True, "breakout": breakout2,
                                 "detail": f"軌道上緣≈{upper_at_last:.2f}　現價＝{last_close:.2f}" + ("　帶量" if vol_confirm else "　量未放大"),
                                 "desc": "股價沿上升軌道緩步上漲，MA20上揚下帶量長紅收盤突破軌道上緣為買點",
                                 "line": {"i1": l1, "p1": ly1 + offset, "slope": slope2},
                                 "line2": {"i1": l1, "p1": ly1, "slope": slope2}})
                added = True
    if not added:
        results.append({"id": id_, "name": name, "formed": False, "breakout": False,
                         "detail": "尚未偵測到符合結構", "desc": "股價沿上升軌道緩步上漲，MA20上揚下帶量長紅收盤突破軌道上緣為買點"})

    # (9) 突破飆股大量黑K最高點
    id_, name = "blackk", "突破飆股大量黑K最高點"
    added = False
    lookback, confirm_window = 10, 3
    floor3 = max(0, len(data) - 1 - lookback)
    candidates = [i for i in range(floor3, last)
                  if data[i]["close"] < data[i]["open"] and data[i]["vm20"] and data[i]["volume"] / data[i]["vm20"] >= 1.6]
    if candidates:
        bk_idx = candidates[-1]
        bk_high = data[bk_idx]["high"]
        within_window = 0 <= (last - bk_idx) <= confirm_window
        if within_window:
            is_red3 = data[last]["close"] > data[last]["open"]
            ma20up3 = _ma_slope_up(data, "ma20", 10, last)
            breakout3 = last_close > bk_high and is_red3 and vol_confirm and ma20up3
            results.append({"id": id_, "name": name, "formed": True, "breakout": breakout3,
                             "detail": f"大量黑K高點＝{bk_high:.2f}　現價＝{last_close:.2f}" + ("　帶量" if vol_confirm else "　量未放大"),
                             "desc": "飆股急漲後出現大量黑K回檔，3日內帶量長紅突破其最高點為買點",
                             "line": {"i1": bk_idx, "p1": bk_high, "slope": 0}})
            added = True
    if not added:
        results.append({"id": id_, "name": name, "formed": False, "breakout": False,
                         "detail": "尚未偵測到符合結構", "desc": "飆股急漲後出現大量黑K回檔，3日內帶量長紅突破其最高點為買點"})

    # (10) K線橫盤的突破：三天以上(含首日)收盤未突破/跌破首日K線高低點，帶量突破首日高點為買點
    id_, name = "kbp", "K線橫盤的突破"
    added = False
    if len(data) >= 4:
        anchor_idx = last - 3  # 最小需求：3天(含首日)
        anchor_bar = data[anchor_idx]
        min_ok = all(anchor_bar["low"] <= data[j]["close"] <= anchor_bar["high"] for j in range(anchor_idx + 1, last))
        if min_ok:
            # 課本圖上的橫盤區間常常不只3天——只要收盤價持續守在「首日K線」的高低範圍內，
            # 就往前延伸找到最早、仍然成立的起點，涵蓋較長的橫盤整理段。
            max_lookback = 20
            candidate = anchor_idx - 1
            while candidate >= 0 and (last - candidate) <= max_lookback:
                in_range = all(data[candidate]["low"] <= data[k]["close"] <= data[candidate]["high"]
                                for k in range(candidate + 1, last))
                if not in_range:
                    break
                anchor_idx = candidate
                candidate -= 1
            anchor = data[anchor_idx]
            is_red_k = data[last]["close"] > data[last]["open"]
            breakout_k = last_close > anchor["high"] and is_red_k and vol_confirm
            results.append({"id": id_, "name": name, "formed": True, "breakout": breakout_k,
                             "detail": f"首日K線高點＝{anchor['high']:.2f}　整理天數＝{last - anchor_idx}天　現價＝{last_close:.2f}" + ("　(帶量)" if vol_confirm else "　(量未放大)"),
                             "desc": "三天以上(含首日)收盤未突破首日K線高低點，帶量紅K突破首日高點為買點",
                             "line": {"i1": anchor_idx, "p1": anchor["high"], "slope": 0}})
            added = True
    if not added:
        results.append({"id": id_, "name": name, "formed": False, "breakout": False,
                         "detail": "尚未偵測到符合結構", "desc": "三天以上(含首日)收盤未突破首日K線高低點，帶量紅K突破首日高點為買點"})

    # (11) 高檔母子懷抱：中長紅K(母)＋隔日不過高不破低的黑K/變盤線(子)，次日確認轉折向下
    # 依課程講義：屬於2根K線構成，上漲高檔出現中長紅，次日出現不過高也不破低的黑K線，
    # 中長紅K線稱為母線，次日K線稱為子線；代表多空力量突然拉鋸，多頭上漲力道減弱。
    id_, name = "harami_bear", "母子懷抱(高檔)"
    added = False
    if len(data) >= 3:
        mother, child, confirm_day = data[last - 2], data[last - 1], data[last]
        mother_body_pct = abs(mother["close"] - mother["open"]) / mother["close"]
        mother_is_red = mother["close"] > mother["open"]
        mother_is_med_long = mother_body_pct >= 0.035
        child_contained = child["high"] <= mother["high"] and child["low"] >= mother["low"]
        child_smaller = abs(child["close"] - child["open"]) < abs(mother["close"] - mother["open"]) * 0.6
        at_high = mother["close"] >= data[last - 3]["close"] if last - 3 >= 0 else True
        if mother_is_red and mother_is_med_long and child_contained and child_smaller and at_high:
            breakout_h = confirm_day["close"] < child["close"]
            results.append({"id": id_, "name": name, "formed": True, "breakout": breakout_h,
                             "detail": f"母K(中長紅)高點＝{mother['high']:.2f}　子K收於母K範圍內　" + ("次日已確認轉折向下" if breakout_h else "等待次日確認轉折向下"),
                             "desc": "上漲高檔出現中長紅K，次日不過高不破低的黑K線為母子懷抱，多頭上漲力道轉弱",
                             "marker": {"idx": last - 1, "price": mother["high"], "dir": "up", "label": "母子懷抱"}})
            added = True
    if not added:
        results.append({"id": id_, "name": name, "formed": False, "breakout": False,
                         "detail": "尚未偵測到符合結構", "desc": "上漲高檔出現中長紅K，次日不過高不破低的黑K線為母子懷抱，多頭上漲力道轉弱"})

    # (12) 低檔母子懷抱：中長黑K(母)＋隔日不過高不破低的紅K/變盤線(子)，次日確認轉折向上
    id_, name = "harami_bull", "母子懷抱(低檔)"
    added = False
    if len(data) >= 3:
        mother, child, confirm_day = data[last - 2], data[last - 1], data[last]
        mother_body_pct = abs(mother["close"] - mother["open"]) / mother["close"]
        mother_is_black = mother["close"] < mother["open"]
        mother_is_med_long = mother_body_pct >= 0.035
        child_contained = child["high"] <= mother["high"] and child["low"] >= mother["low"]
        child_smaller = abs(child["close"] - child["open"]) < abs(mother["close"] - mother["open"]) * 0.6
        at_low = mother["close"] <= data[last - 3]["close"] if last - 3 >= 0 else True
        if mother_is_black and mother_is_med_long and child_contained and child_smaller and at_low:
            breakout_h2 = confirm_day["close"] > child["close"]
            results.append({"id": id_, "name": name, "formed": True, "breakout": breakout_h2,
                             "detail": f"母K(中長黑)低點＝{mother['low']:.2f}　子K收於母K範圍內　" + ("次日已確認轉折向上" if breakout_h2 else "等待次日確認轉折向上"),
                             "desc": "下跌低檔出現中長黑K，次日不過高不破低的紅K線為母子懷抱，空頭下跌力道轉弱",
                             "marker": {"idx": last - 1, "price": mother["low"], "dir": "down", "label": "母子懷抱"}})
            added = True
    if not added:
        results.append({"id": id_, "name": name, "formed": False, "breakout": False,
                         "detail": "尚未偵測到符合結構", "desc": "下跌低檔出現中長黑K，次日不過高不破低的紅K線為母子懷抱，空頭下跌力道轉弱"})

    # (13) 晨星：下跌中長黑K＋變盤線＋中長紅K，收盤站上首日黑K實體中點為低檔轉折向上
    # 依課程講義：下跌低檔出現左邊中長黑K，右邊中長紅K，中間夾一根「變盤線」，是低檔轉折向上確認的K線組合。
    id_, name = "morning_star", "晨星"
    added = False
    if len(data) >= 3:
        s1, s2, s3 = data[last - 2], data[last - 1], data[last]
        s1_body_pct = abs(s1["close"] - s1["open"]) / s1["close"]
        s1_is_black = s1["close"] < s1["open"]
        s1_med_long = s1_body_pct >= 0.035
        s2_body_pct = abs(s2["close"] - s2["open"]) / s2["close"]
        is_star1 = s2_body_pct < 0.035
        s3_body_pct = abs(s3["close"] - s3["open"]) / s3["close"]
        s3_is_red = s3["close"] > s3["open"]
        s3_med_long = s3_body_pct >= 0.035
        s1_mid = (s1["open"] + s1["close"]) / 2
        closes_above_mid = s3["close"] > s1_mid
        if s1_is_black and s1_med_long and is_star1 and s3_is_red and s3_med_long and closes_above_mid:
            results.append({"id": id_, "name": name, "formed": True, "breakout": True,
                             "detail": f"首日黑K實體中點＝{s1_mid:.2f}　收盤＝{s3['close']:.2f}（已站上中點，轉折確認）",
                             "desc": "下跌出現中長黑K+變盤線+中長紅K，收盤站上首日實體中點為低檔轉折向上訊號",
                             "marker": {"idx": last - 1, "price": s2["low"], "dir": "down", "label": "晨星"}})
            added = True
    if not added:
        results.append({"id": id_, "name": name, "formed": False, "breakout": False,
                         "detail": "尚未偵測到符合結構", "desc": "下跌出現中長黑K+變盤線+中長紅K，收盤站上首日實體中點為低檔轉折向上訊號"})

    # (14) 夜星：上漲中長紅K＋變盤線＋中長黑K，收盤跌破首日紅K實體中點為高檔轉折向下
    # 晨星的鏡像型態（課程講義未獨立列出，依同一套邏輯對稱推導）
    id_, name = "evening_star", "夜星"
    added = False
    if len(data) >= 3:
        e1, e2, e3 = data[last - 2], data[last - 1], data[last]
        e1_body_pct = abs(e1["close"] - e1["open"]) / e1["close"]
        e1_is_red = e1["close"] > e1["open"]
        e1_med_long = e1_body_pct >= 0.035
        e2_body_pct = abs(e2["close"] - e2["open"]) / e2["close"]
        is_star2 = e2_body_pct < 0.035
        e3_body_pct = abs(e3["close"] - e3["open"]) / e3["close"]
        e3_is_black = e3["close"] < e3["open"]
        e3_med_long = e3_body_pct >= 0.035
        e1_mid = (e1["open"] + e1["close"]) / 2
        closes_below_mid = e3["close"] < e1_mid
        if e1_is_red and e1_med_long and is_star2 and e3_is_black and e3_med_long and closes_below_mid:
            results.append({"id": id_, "name": name, "formed": True, "breakout": True,
                             "detail": f"首日紅K實體中點＝{e1_mid:.2f}　收盤＝{e3['close']:.2f}（已跌破中點，轉折確認）",
                             "desc": "上漲出現中長紅K+變盤線+中長黑K，收盤跌破首日實體中點為高檔轉折向下訊號",
                             "marker": {"idx": last - 1, "price": e2["high"], "dir": "up", "label": "夜星"}})
            added = True
    if not added:
        results.append({"id": id_, "name": name, "formed": False, "breakout": False,
                         "detail": "尚未偵測到符合結構", "desc": "上漲出現中長紅K+變盤線+中長黑K，收盤跌破首日實體中點為高檔轉折向下訊號"})

    # (15) 回後買上漲：沿用 checkPullbackBuy() 判斷結果
    if pb:
        pb_formed = pb["allPass"] or pb["requiredPassed"] >= pb["requiredTotal"] - 1
        results.append({
            "id": "pbup", "name": "回後買上漲",
            "formed": pb_formed, "breakout": pb["allPass"],
            "detail": f"必要條件 {pb['requiredPassed']}/{pb['requiredTotal']} 通過" + ("　+成交量增加分" if pb["bonusPassed"] else ""),
            "desc": "趨勢多頭，回檔量縮價穩後，今日紅K放量突破前高為買點",
        })

    # ── 剛突破：與前一交易日比較，突破訊號是「今天才發生」──
    if not skip_just_broke and len(data) > 1:
        prev_data = data[:-1]
        prev_pb = check_pullback_buy(prev_data)
        prev_pt = detect_patterns(prev_data, prev_pb, skip_just_broke=True)
        for i, r in enumerate(results):
            pr = prev_pt["results"][i] if i < len(prev_pt["results"]) else None
            r["justBroke"] = bool(r["breakout"] and (not pr or not pr["breakout"]))
    else:
        for r in results:
            r["justBroke"] = False

    any_breakout = any(r["breakout"] for r in results)
    any_formed = any(r["formed"] for r in results)
    any_just_broke = any(r["justBroke"] for r in results)

    return {"results": results, "anyBreakout": any_breakout, "anyFormed": any_formed, "anyJustBroke": any_just_broke}


# ────────────────────────────────────────────────────────────────
# Plotly 圖表：K線＋均線＋布林通道＋成交量＋MACD＋轉折波
# ────────────────────────────────────────────────────────────────

def _is_finite(x):
    try:
        return x == x and x not in (float("inf"), float("-inf"))
    except TypeError:
        return False


def _is_sane_bar(d):
    """過濾掉資料來源偶爾出現的異常值（例如某天 low 被錯誤回傳為 0 或極小值），
    避免單一根爛資料把轉折波拉出一條不合理的長長尖刺"""
    if not d:
        return False
    high, low, close = d.get("high"), d.get("low"), d.get("close")
    if not (high is not None and high > 0 and low is not None and low > 0 and close is not None and close > 0):
        return False
    import math
    if not (math.isfinite(high) and math.isfinite(low) and math.isfinite(close)):
        return False
    ma5 = d.get("ma5")
    if ma5 is not None and ma5 > 0:
        if low < ma5 * 0.4 or high > ma5 * 2.5:
            return False
    return True


def build_zigzag(data):
    """朱家泓「短線轉折波」畫法（依課程圖表2-1-3／2-1-4）：
    用收盤價與5日均線的穿越關係取高低點——
      收盤價「跌破」5日均線 → 取這段上漲過程的「高點」為一個轉折點
      收盤價「突破」5日均線 → 取這段下跌過程的「低點」為一個轉折點
    再把這些高低點依序連接成鋸齒狀的轉折波。"""
    n = len(data)
    # 找到第一個 MA5 已經有值、且資料正常的位置（前4根K棒沒有5日均線可比較）
    start = 0
    while start < n and (data[start]["ma5"] is None or not _is_sane_bar(data[start])):
        start += 1
    if start >= n - 1:
        return []

    state = "above" if data[start]["close"] >= data[start]["ma5"] else "below"  # 目前收盤在5日均線上方或下方
    points = [{
        "idx": start,
        "price": data[start]["low"] if state == "above" else data[start]["high"],
        "type": "L" if state == "above" else "H",
    }]
    extreme_idx = start
    extreme_price = data[start]["high"] if state == "above" else data[start]["low"]

    for i in range(start + 1, n):
        d = data[i]
        if d["ma5"] is None or not _is_sane_bar(d):
            continue
        if state == "above":
            # 收盤在5日均線之上，持續追蹤這段上漲的最高點
            if d["high"] > extreme_price:
                extreme_price, extreme_idx = d["high"], i
            if d["close"] < d["ma5"]:
                # 收盤跌破5日均線 → 確認剛才追蹤到的高點為一個轉折高點
                points.append({"idx": extreme_idx, "price": extreme_price, "type": "H"})
                state = "below"
                extreme_price, extreme_idx = d["low"], i
        else:
            # 收盤在5日均線之下，持續追蹤這段下跌的最低點
            if d["low"] < extreme_price:
                extreme_price, extreme_idx = d["low"], i
            if d["close"] > d["ma5"]:
                # 收盤突破5日均線 → 確認剛才追蹤到的低點為一個轉折低點
                points.append({"idx": extreme_idx, "price": extreme_price, "type": "L"})
                state = "above"
                extreme_price, extreme_idx = d["high"], i

    # 收尾：把目前仍在追蹤中的高/低點畫出來，再接到最新一根K棒的收盤價，確保線一定連到最新資料
    last_idx = n - 1
    if extreme_idx != last_idx:
        points.append({"idx": extreme_idx, "price": extreme_price, "type": "H" if state == "above" else "L"})
        points.append({"idx": last_idx, "price": data[last_idx]["close"], "type": "L" if state == "above" else "H"})
    else:
        points.append({"idx": extreme_idx, "price": data[last_idx]["close"], "type": "H" if state == "above" else "L"})
    return points


def draw_chart(data, name, pt=None):
    # 圖表顯示範圍改用完整的 data（已由「分析天數」在抓資料時決定範圍），
    # 不再寫死只看最近120根K棒，讓滑桿調整能真正反映在圖表上。
    raw_tail = data
    # 過濾掉資料異常的K棒（開高低收有任一項是 0、負值或非數字），
    # 避免圖表出現「沒有K棒的空白位置」卻仍有轉折波或均線的線硬穿過去
    tail = [d for d in raw_tail if d.get("open", 0) > 0 and d.get("high", 0) > 0
            and d.get("low", 0) > 0 and d.get("close", 0) > 0
            and all(_is_finite(d[k]) for k in ("open", "high", "low", "close"))]
    dates = [d["date"] for d in tail]
    vol_colors = ["#ef5350" if d["close"] >= d["open"] else "#26a69a" for d in tail]
    hist_colors = ["#ef5350" if (d["macdHist"] or 0) >= 0 else "#26a69a" for d in tail]

    zz = build_zigzag(tail)
    zz_x = [tail[p["idx"]]["date"] for p in zz]
    zz_y = [p["price"] for p in zz]

    # 每個「型態確認」如果已成形，就把輔助線（頸線/壓力線/切線/軌道線）畫在圖上，
    # 型態成形但未突破 → 虛線；已經突破 → 實線＋★標註，方便直接在圖上對照型態辨識依據
    PATTERN_LINE_STYLE = {
        "hs": {"color": "#ffd54f", "label": "頭肩底頸線"},
        "chs": {"color": "#ff8a65", "label": "複式頭肩底頸線"},
        "nb": {"color": "#81c784", "label": "N字底壓力"},
        "tb": {"color": "#4dd0e1", "label": "三重底壓力"},
        "rb": {"color": "#64b5f6", "label": "圓弧底壓力"},
        "fb": {"color": "#ba68c8", "label": "一字底整理區間高點"},
        "abc": {"color": "#ff2ecc", "label": "ABC下降切線起點"},
        "channel": {"color": "#ffa726", "label": "上升軌道線"},
        "blackk": {"color": "#e57373", "label": "大量黑K高點"},
        "kbp": {"color": "rgba(255,255,255,.85)", "label": "K線橫盤首日高點"},
    }
    # 母子懷抱／晨星／夜星屬於2~3根K線的短線反轉訊號，沒有持續延伸的支撐/壓力線可畫，
    # 改用箭頭標註直接指到型態發生的那根（或那兩根）K棒位置，方便在圖上直接辨識。
    PATTERN_MARKER_STYLE = {
        "harami_bear": {"color": "#ff8a65"},
        "harami_bull": {"color": "#81c784"},
        "morning_star": {"color": "#4dd0e1"},
        "evening_star": {"color": "#ff5252"},
    }
    shapes, annotations = [], []
    last_idx = len(tail) - 1
    if pt and pt.get("results"):
        for p in pt["results"]:
            if not p.get("formed"):
                continue
            style = PATTERN_LINE_STYLE.get(p["id"])
            if style:
                # 上升軌道線另外多畫一條下緣支撐線（line2），其餘型態只有一條輔助線（line）
                for idx, ln in enumerate((p.get("line"), p.get("line2"))):
                    if not ln or ln["i1"] < 0 or ln["i1"] >= len(tail):
                        continue
                    line_end_price = ln["p1"] + ln["slope"] * (last_idx - ln["i1"])
                    shapes.append(dict(
                        type="line", xref="x", yref="y",
                        x0=tail[ln["i1"]]["date"], y0=ln["p1"],
                        x1=tail[last_idx]["date"], y1=line_end_price,
                        line=dict(color=style["color"], width=2 if idx == 0 else 1.3,
                                   dash="solid" if p["breakout"] else "dash"),
                    ))
                    if idx == 0:
                        annotations.append(dict(
                            x=tail[ln["i1"]]["date"], y=ln["p1"], xref="x", yref="y",
                            text=style["label"], showarrow=True, arrowhead=2, arrowcolor=style["color"],
                            font=dict(color=style["color"], size=10), ax=-10, ay=-30,
                        ))
                        if p["breakout"]:
                            annotations.append(dict(
                                x=tail[last_idx]["date"], y=tail[last_idx]["close"], xref="x", yref="y",
                                text="★ 突破" + p["name"], showarrow=True, arrowhead=2, arrowcolor=style["color"],
                                font=dict(color=style["color"], size=11), ax=10, ay=-35,
                            ))
                continue
            m_style = PATTERN_MARKER_STYLE.get(p["id"])
            marker = p.get("marker")
            if m_style and marker and 0 <= marker["idx"] < len(tail):
                annotations.append(dict(
                    x=tail[marker["idx"]]["date"], y=marker["price"], xref="x", yref="y",
                    text=("★ " if p["breakout"] else "") + marker["label"],
                    showarrow=True, arrowhead=2, arrowcolor=m_style["color"],
                    font=dict(color=m_style["color"], size=11),
                    ax=0, ay=-32 if marker["dir"] == "up" else 32,
                ))

    fig = make_subplots(rows=3, cols=1, shared_xaxes=True, row_heights=[0.56, 0.22, 0.22], vertical_spacing=0.03)

    fig.add_trace(go.Candlestick(
        x=dates, open=[d["open"] for d in tail], high=[d["high"] for d in tail],
        low=[d["low"] for d in tail], close=[d["close"] for d in tail], name="K線",
        increasing=dict(line=dict(color="#ef5350"), fillcolor="#ef5350"),
        decreasing=dict(line=dict(color="#26a69a"), fillcolor="#26a69a"),
    ), row=1, col=1)

    for key, color, label in [("ma5", "#ffeb3b", "MA5"), ("ma10", "#ff9800", "MA10"),
                               ("ma20", "#2196f3", "MA20"), ("ma60", "#9c27b0", "MA60")]:
        fig.add_trace(go.Scatter(x=dates, y=[d[key] for d in tail], name=label,
                                  line=dict(color=color, width=1.2)), row=1, col=1)

    fig.add_trace(go.Scatter(x=zz_x, y=zz_y, name="轉折波", mode="lines+markers",
                              line=dict(color="#00e5ff", width=1.8),
                              marker=dict(size=5, color="#00e5ff")), row=1, col=1)

    fig.add_trace(go.Scatter(x=dates, y=[d["bbU"] for d in tail], name="BB上軌",
                              line=dict(color="rgba(100,200,255,.4)", width=1, dash="dot"),
                              showlegend=False), row=1, col=1)
    fig.add_trace(go.Scatter(x=dates, y=[d["bbL"] for d in tail], name="BB下軌",
                              line=dict(color="rgba(100,200,255,.4)", width=1, dash="dot"),
                              fill="tonexty", fillcolor="rgba(100,200,255,.04)", showlegend=False), row=1, col=1)

    fig.add_trace(go.Bar(x=dates, y=[d["volume"] for d in tail], marker_color=vol_colors,
                          name="成交量", opacity=0.7), row=2, col=1)
    fig.add_trace(go.Scatter(x=dates, y=[d["vm20"] for d in tail], name="量MA20",
                              line=dict(color="#ff9800", width=1.5)), row=2, col=1)

    fig.add_trace(go.Bar(x=dates, y=[d["macdHist"] for d in tail], marker_color=hist_colors,
                          name="MACD柱", opacity=0.8), row=3, col=1)
    fig.add_trace(go.Scatter(x=dates, y=[d["macd"] for d in tail], name="MACD",
                              line=dict(color="#2196f3", width=1.5)), row=3, col=1)
    fig.add_trace(go.Scatter(x=dates, y=[d["macdSig"] for d in tail], name="Signal",
                              line=dict(color="#ff9800", width=1.5)), row=3, col=1)

    fig.update_layout(
        title=dict(text=f"{name} 技術分析圖", font=dict(color="#fff", size=14)),
        template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(15,20,35,1)",
        height=680, margin=dict(l=50, r=15, t=45, b=25),
        shapes=shapes, annotations=annotations,
        xaxis=dict(rangeslider=dict(visible=False), gridcolor="rgba(255,255,255,.04)"),
        legend=dict(orientation="h", y=1.06, bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
        showlegend=True,
    )
    fig.update_yaxes(gridcolor="rgba(255,255,255,.04)")
    return fig


# ────────────────────────────────────────────────────────────────
# AI 智能綜合分析（呼叫 OpenAI API）
# ────────────────────────────────────────────────────────────────

def build_analysis_prompt(r):
    last = r["data"][-1]
    prev = r["data"][-2] if len(r["data"]) >= 2 else last
    chg = last["close"] - prev["close"]
    chgp = (chg / prev["close"] * 100) if prev["close"] else 0

    lines = []
    lines.append(f"股票：{r['stockId']} {r['name']}")
    lines.append(f"資料日期：{last['date']}　收盤：{last['close']}　漲跌：{'+' if chg >= 0 else ''}{chg:.2f} ({'+' if chgp >= 0 else ''}{chgp:.2f}%)")
    vol_ratio_txt = f"{(last['volume'] / last['vm20']):.2f}" if last["vm20"] else "N/A"
    lines.append(f"成交量：{last['volume']}　量比(vs MA20量)：{vol_ratio_txt}x")
    if last.get("macd") is not None and last.get("macdSig") is not None:
        bias = "（偏多）" if last["macd"] > last["macdSig"] else "（偏空）"
        lines.append(f"MACD：DIF={last['macd']:.2f}　Signal={last['macdSig']:.2f}　柱狀={last['macdHist']:.2f}{bias}")
    if last.get("bbU") is not None and last.get("bbL") is not None:
        bb_w = last["bbU"] - last["bbL"]
        bb_p = ((last["close"] - last["bbL"]) / bb_w * 100) if bb_w else 50
        lines.append(f"布林通道：上軌={last['bbU']:.2f}　下軌={last['bbL']:.2f}　股價位置={bb_p:.0f}%")
    lines.append("")
    lines.append(f"【DMI多方力道評分】總分 {r['total']}/100")
    dm = r["dm"]
    lines.append(f"+DI={dm['plusDI']:.1f}　-DI={dm['minusDI']:.1f}　ADX={dm['adx']:.1f}　ADXR={dm['adxr']:.1f}"
                 if dm["plusDI"] is not None and dm["adx"] is not None and dm["adxr"] is not None
                 else "DMI資料不足")
    lines.append(f"方向性 {dm['diPts']:.1f}/50、趨勢強度 {dm['adxPts']:.1f}/30、趨勢動能 {dm['adxrPts']:.1f}/20")
    all_sigs = dm["sigs"]
    bull_sigs = [s[0] for s in all_sigs if s[1] == "bull"]
    bear_sigs = [s[0] for s in all_sigs if s[1] == "bear"]
    if bull_sigs:
        lines.append("多頭訊號：" + "、".join(bull_sigs))
    if bear_sigs:
        lines.append("空頭訊號：" + "、".join(bear_sigs))
    lines.append("")
    lines.append(f"【回後買上漲 8條件核對】必要條件通過 {r['pb']['requiredPassed']}/{r['pb']['requiredTotal']}" + ("（全數通過）" if r["pb"]["allPass"] else ""))
    lines.append("")
    lines.append("【型態確認，15種進場型態】")
    for p in r["pt"]["results"]:
        status = "🔥剛突破（較前一交易日新增）" if p["justBroke"] else ("✅已突破" if p["breakout"] else ("🕒成形中未突破" if p["formed"] else "－未偵測到"))
        lines.append(f"・{p['name']}：{status}" + (f"（{p['detail']}）" if p.get("detail") else ""))
    return "\n".join(lines)


def run_ai_analysis(r, api_key, model):
    prompt = build_analysis_prompt(r)
    try:
        res = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
            json={
                "model": model,
                "temperature": 0.4,
                "max_tokens": 900,
                "messages": [
                    {"role": "system", "content": "你是一位精通台股技術分析的資深操盤手，熟悉DMI趨向指標（+DI／-DI／ADX／ADXR）多方力道評分方法論，以及朱家泓《技術分析全攻略》的回後買上漲、頭肩底等進場型態判斷。請根據使用者提供的個股技術數據摘要，用繁體中文給出：1) 整體技術面研判（3-4句） 2) 進場時機與風險提示 3) 綜合建議（積極做多／可考慮／觀望／不建議）。語氣專業、精簡、避免空泛用詞，並提醒這僅為技術面參考，非投資建議。"},
                    {"role": "user", "content": prompt},
                ],
            },
            timeout=60,
        )
        j = res.json()
        if not res.ok:
            msg = (j.get("error") or {}).get("message") or f"HTTP {res.status_code}"
            raise RuntimeError(msg)
        content = ((j.get("choices") or [{}])[0].get("message") or {}).get("content")
        return content or "（AI 未回傳有效內容）"
    except Exception as e:
        return f"❌ AI 分析失敗：{e}\n請確認 API Key 是否正確、額度是否足夠，或稍後再試。"


# ────────────────────────────────────────────────────────────────
# Streamlit UI
# ────────────────────────────────────────────────────────────────

if "stock_text_input" not in st.session_state:
    st.session_state["stock_text_input"] = "\n".join(TWSE_LIST)
if "active_list" not in st.session_state:
    st.session_state.active_list = "twse"
if "custom_lists" not in st.session_state:
    st.session_state.custom_lists = load_custom_lists()

st.title("📊 技術分析全攻略 · 個股評分分析系統")
st.caption("個股評分改為DMI趨向指標（+DI／-DI／ADX／ADXR）多方力道評分；回後買上漲、15種進場型態辨識仍沿用朱家泓《技術分析全攻略》方法論")


def fetch_taiwan_market_snapshot(token: str):
    """今日全市場快照：改用 FinMind 自己的 API（跟本工具其他功能同一個資料來源），
    不帶 data_id 直接查全市場當天資料，此為 Backer/Sponsor 方案才能使用的用法
    （免費方案只能查單一股票）。流程：①先測試 Token 是否有效 → ②再查全市場當天資料。
    「漲幅前100」「成交量前100」共用同一份快照，只是排序依據不同，避免重複打兩次API。
    回傳 (results, used_date, error_msg)，error_msg 有值代表失敗原因；results 每筆含 id/name/pct/volume。"""
    # ① 測試 Token 是否有效
    try:
        test_json = api_fetch(f"{FINMIND_BASE}?dataset=TaiwanStockInfo&token={token}")
    except Exception as e:
        return [], None, f"❌ 無法連線至 FinMind API：{e}\n請確認網路連線是否正常。"
    if test_json.get("status") != 200:
        msg = test_json.get("msg") or f"status={test_json.get('status')}"
        return [], None, f"❌ Token 無效或 API 錯誤：{msg}\n請確認 FinMind Token 是否正確。"

    # ② 查全市場當天資料（往前找最近一個有資料的交易日，最多找7天，逐日獨立重試）
    price_json, used_date, last_day_err = None, None, ""
    for back in range(7):
        date_str = (datetime.today() - timedelta(days=back)).strftime("%Y-%m-%d")
        try:
            j = api_fetch(f"{FINMIND_BASE}?dataset=TaiwanStockPrice&start_date={date_str}&token={token}")
        except Exception as e:
            last_day_err = str(e)
            continue
        if j.get("status") == 200 and j.get("data"):
            price_json, used_date = j, date_str
            break
        last_day_err = j.get("msg") or f"status={j.get('status')}"
    if price_json is None:
        return [], None, (
            f"❌ 近7天皆無法取得全市場資料\n錯誤：{last_day_err}\n\n"
            "這個功能需要 FinMind「Backer」或「Sponsor」付費方案才能查詢全市場單日資料，"
            "免費方案只能查詢單一股票。請確認您的帳號方案，或改用「上市清單／上櫃清單」搭配「批次分析」。"
        )

    # ③ 股票名稱對照（用①已經拿到的 TaiwanStockInfo 結果，不用再查一次）
    name_map = {}
    for item in test_json.get("data", []):
        sid = item.get("stock_id")
        sname = item.get("stock_name")
        if sid and sname and sid not in name_map:
            name_map[sid] = sname

    # 只保留出現在本工具「上市清單」「上櫃清單」裡的代碼——這兩份清單本來就是精選過的
    # 一般個股代碼，直接拿來當白名單，能同時排除ETF、權證、TDR、公司債等各種非個股商品，
    # 比單純判斷代碼開頭字元（例如「00」開頭）可靠很多。
    valid_stock_set = set(TWSE_LIST) | set(TPEX_LIST)

    results = []
    for r in price_json["data"]:
        sid = r.get("stock_id")
        if not sid or sid not in valid_stock_set:
            continue
        try:
            close = float(r.get("close"))
            spread = float(r.get("spread"))
        except (TypeError, ValueError):
            continue
        if close <= 0:
            continue
        prev_close = close - spread
        if prev_close <= 0:
            continue
        try:
            volume = float(r.get("Trading_Volume"))
        except (TypeError, ValueError):
            volume = 0.0
        results.append({"id": sid, "name": name_map.get(sid, sid), "pct": (spread / prev_close) * 100, "volume": volume})

    if not results:
        return [], None, "取得的全市場資料是空的，請稍後再試一次。"

    return results, used_date, None


def fetch_top100_gainers(token: str):
    results, used_date, err_msg = fetch_taiwan_market_snapshot(token)
    if err_msg:
        return [], None, err_msg
    top100 = sorted(results, key=lambda r: r["pct"], reverse=True)[:100]
    return top100, used_date, None


def fetch_top100_volume(token: str):
    results, used_date, err_msg = fetch_taiwan_market_snapshot(token)
    if err_msg:
        return [], None, err_msg
    top100 = sorted(results, key=lambda r: r["volume"], reverse=True)[:100]
    return top100, used_date, None


with st.sidebar:
    st.header("📊 技術分析全攻略")
    st.caption("朱家泓方法論 · 個股評分系統")

    api_token = st.text_input("FinMind API Token", type="password", placeholder="輸入您的Token")

    st.subheader("批次股票代號")
    c1, c2, c3 = st.columns(3)
    if c1.button("上市清單", use_container_width=True,
                  type="primary" if st.session_state.active_list == "twse" else "secondary"):
        st.session_state["stock_text_input"] = "\n".join(TWSE_LIST)
        st.session_state.active_list = "twse"
        st.rerun()
    if c2.button("上櫃清單", use_container_width=True,
                  type="primary" if st.session_state.active_list == "tpex" else "secondary"):
        st.session_state["stock_text_input"] = "\n".join(TPEX_LIST)
        st.session_state.active_list = "tpex"
        st.rerun()
    if c3.button("0050成分股", use_container_width=True,
                  type="primary" if st.session_state.active_list == "0050" else "secondary"):
        st.session_state["stock_text_input"] = "\n".join(TW0050_LIST)
        st.session_state.active_list = "0050"
        st.rerun()

    # 我的清單／清單1／清單2／清單3：每列一個選取按鈕＋一個 × 清除按鈕
    for key in CUSTOM_LIST_KEYS:
        col_sel, col_clear = st.columns([5, 1])
        items = st.session_state.custom_lists.get(key, [])
        label = CUSTOM_LIST_LABELS[key] + (f"（{len(items)}）" if items else "")
        if col_sel.button(label, use_container_width=True, key=f"sel_{key}",
                           type="primary" if st.session_state.active_list == key else "secondary"):
            st.session_state["stock_text_input"] = "\n".join(items)
            st.session_state.active_list = key
            st.rerun()
        if col_clear.button("×", use_container_width=True, key=f"clear_{key}",
                             help=f"清除「{CUSTOM_LIST_LABELS[key]}」"):
            st.session_state.custom_lists[key] = []
            save_custom_lists(st.session_state.custom_lists)
            if st.session_state.active_list == key:
                st.session_state["stock_text_input"] = ""
            st.rerun()

    # ── 漲幅前100／成交量前100：實際抓資料與設定 stock_text_input 的邏輯，
    # 必須在 text_area 元件實例化「之前」執行，否則會觸發 StreamlitAPIException
    # （Streamlit 不允許在 widget 已經渲染後才設定該 widget key 對應的 session_state）。
    # 所以按鈕本身仍畫在 text_area 之後（視覺順序不變），按下後只設一個旗標並 rerun，
    # 實際抓資料的邏輯則挪到這裡、text_area 之前執行。
    if st.session_state.pop("_trigger_top100_gainers", False):
        if not api_token:
            st.session_state["top100_info"] = "❌ 請先輸入 FinMind API Token 才能抓取今日漲幅排行（需 Backer/Sponsor 方案）"
        else:
            with st.spinner("① 測試 FinMind API Token 中…"):
                top100, used_date, err_msg = fetch_top100_gainers(api_token)
            if top100:
                st.session_state["stock_text_input"] = "\n".join(r["id"] for r in top100)
                st.session_state.active_list = "top100"
                st.session_state["top100_info"] = (
                    f"✅ 已套入 {used_date} 漲幅前{len(top100)}"
                    f"（最高：{top100[0]['id']} {top100[0]['name']} +{top100[0]['pct']:.2f}%），正在自動開始批次分析…"
                )
                # 套入清單成功後，標記在下一次 rerun 時自動接著跑批次分析，不用使用者再按一次「批次分析」
                st.session_state["_run_after_top100"] = True
            else:
                st.session_state["top100_info"] = err_msg or "❌ 未能取得今日行情資料，請稍後再試"

    if st.session_state.pop("_trigger_volume100", False):
        if not api_token:
            st.session_state["top100_info"] = "❌ 請先輸入 FinMind API Token 才能抓取今日成交量排行（需 Backer/Sponsor 方案）"
        else:
            with st.spinner("① 測試 FinMind API Token 中…"):
                vol100, used_date, err_msg = fetch_top100_volume(api_token)
            if vol100:
                st.session_state["stock_text_input"] = "\n".join(r["id"] for r in vol100)
                st.session_state.active_list = "volume100"
                st.session_state["top100_info"] = (
                    f"✅ 已套入 {used_date} 成交量前{len(vol100)}"
                    f"（最高：{vol100[0]['id']} {vol100[0]['name']} 量={int(vol100[0]['volume']):,} 股），正在自動開始批次分析…"
                )
                st.session_state["_run_after_top100"] = True
            else:
                st.session_state["top100_info"] = err_msg or "❌ 未能取得今日行情資料，請稍後再試"

    if st.session_state.get("top100_info"):
        (st.success if st.session_state["top100_info"].startswith("✅") else st.error)(
            st.session_state["top100_info"]
        )

    stock_text = st.text_area("每行一個，或逗號分隔", height=180, key="stock_text_input")

    if st.button("💾 更新我的清單", use_container_width=True):
        stocks = parse_stock_tokens(stock_text)
        if stocks:
            st.session_state.custom_lists["my"] = stocks
            save_custom_lists(st.session_state.custom_lists)
            st.success(f"已更新我的清單（{len(stocks)}檔）")
        else:
            st.warning("批次股票代號目前是空的，沒有可儲存的內容")

    # 存為清單1／清單2／清單3：三個各自獨立的按鈕，直接存進指定槽位（不再是舊版
    # 「自動找空槽」的邏輯）。Streamlit沒有原生confirm彈窗，所以跟清單的×清除鈕一樣，
    # 按下就直接覆蓋，不會另外跳確認。
    save_cols = st.columns(3)
    for idx, key in enumerate(("my1", "my2", "my3")):
        if save_cols[idx].button(f"➕ 存為清單{idx+1}", use_container_width=True, key=f"save_{key}"):
            stocks = parse_stock_tokens(stock_text)
            if stocks:
                st.session_state.custom_lists[key] = stocks
                save_custom_lists(st.session_state.custom_lists)
                st.success(f"已存入{CUSTOM_LIST_LABELS[key]}（{len(stocks)}檔）")
            else:
                st.warning("批次股票代號目前是空的，沒有可儲存的內容")

    days = st.slider("分析天數", min_value=90, max_value=365, value=180, step=30)
    use_realtime = st.checkbox("🔴 加入盤中即時股價（需FinMind sponsor會員，非sponsor會自動略過）", value=True)

    st.caption("以下每勾一項，批次分析都會為每檔股票多打1次API，股票數多時會明顯變慢：")
    show_pe = st.checkbox("📐 近3年P/E區間", value=False)
    show_rev = st.checkbox("📈 近3月營收YoY/MoM", value=False)
    show_pxyoy = st.checkbox("💹 近3月均價YoY（可獨立勾選；若同時勾營收，月份會對齊營收那組）", value=False)
    show_inst = st.checkbox("🏦 三大法人買賣超（近3月，逐月加總）", value=False)

    run_clicked = st.button("🔍 批次分析", type="primary", use_container_width=True)

    with st.expander("📸 全市場快照記錄（建立回測用歷史資料庫）"):
        stats = get_snapshot_stats()
        if stats and stats["total_rows"]:
            st.caption(
                f"目前資料庫：共 {stats['total_rows']:,} 筆快照、"
                f"{stats['distinct_dates']} 個交易日、{stats['distinct_stocks']} 檔股票　"
                f"（{stats['earliest_date']} ～ {stats['latest_date']}）"
            )
        else:
            st.caption("目前資料庫還是空的，尚未記錄過任何快照。")
        st.caption(
            "固定使用「上市清單」（約1,100+檔）當樣本池，不是目前輸入框裡的清單——"
            "樣本數愈大，之後回測才有統計意義。只記錄核心DMI/布林通道/MACD訊號，"
            "不含P/E、營收、法人這些額外欄位。全市場掃描預估需要10-20分鐘，"
            "過程中請勿切換分頁或關閉視窗。"
        )
        snapshot_clicked = st.button("📸 執行全市場快照掃描並寫入資料庫", use_container_width=True)
        if snapshot_clicked:
            if not api_token:
                st.error("請輸入 FinMind API Token")
            else:
                run_full_market_snapshot(api_token)

    with st.expander("🔬 歷史回測分析（過去3個月，事後驗證＋參數優化）"):
        st.caption(
            "回溯過去3個月，用「上市清單」重新計算每個交易日『當時』的DMI/多方力道評分"
            "（只用當天以前的資料，沒有偷看未來），對照5/10/20個交易日後的實際報酬，"
            "驗證現有評分公式準不準，並試算不同參數組合的效果。全市場規模預估需要"
            "15-25分鐘，過程中請勿切換分頁或關閉視窗。"
        )
        backtest_clicked = st.button("🔬 執行歷史回測（過去3個月）", use_container_width=True)
        if backtest_clicked:
            if not api_token:
                st.error("請輸入 FinMind API Token")
            else:
                run_historical_backtest(api_token, months_back=3)

        bt_df = load_backtest_df()
        if bt_df is not None and not bt_df.empty:
            st.markdown(
                f"**目前累積 {len(bt_df):,} 筆評估紀錄**"
                f"（{bt_df['eval_date'].min()} ～ {bt_df['eval_date'].max()}）"
            )

            st.markdown("##### 📊 評分區間 vs 實際報酬")
            st.dataframe(analyze_score_buckets(bt_df), hide_index=True, use_container_width=True)

            st.markdown("##### 🚀 「強勢突破盤」標記 vs 實際報酬")
            st.dataframe(analyze_tag_hitrate(bt_df, "is_breakout", "強勢突破盤"), hide_index=True, use_container_width=True)

            st.markdown("##### 🎯 「跌深反彈盤」標記 vs 實際報酬")
            st.dataframe(analyze_tag_hitrate(bt_df, "is_pullback_rebound", "跌深反彈盤"), hide_index=True, use_container_width=True)

            st.markdown("##### 🎛️ 參數網格搜尋")
            st.caption(
                "⚠️ 這是在已收集的歷史資料上找『表現較好』的參數組合，樣本數有限時容易"
                "過度適配——建議當作方向參考，人工確認合理後再手動調整正式評分公式，"
                "不要照單全收直接套用。"
            )
            horizon_choice = st.selectbox("優化目標天數", BACKTEST_HORIZONS, index=1, key="grid_horizon")
            grid_df = grid_search_params(bt_df, target_horizon=horizon_choice)
            if not grid_df.empty:
                st.dataframe(grid_df, hide_index=True, use_container_width=True)
            else:
                st.caption("資料量還不夠做網格搜尋分析（需要至少20筆訊號才會列入單一組合）。")
        else:
            st.caption("尚未執行過歷史回測，點上面的按鈕開始。")

    top100_clicked = st.button("🔥 漲幅前100分析", use_container_width=True)
    if top100_clicked:
        st.session_state["_trigger_top100_gainers"] = True
        st.rerun()

    volume100_clicked = st.button("📊 成交量前100分析", use_container_width=True)
    if volume100_clicked:
        st.session_state["_trigger_volume100"] = True
        st.rerun()

    st.divider()
    openai_key = st.text_input("OpenAI API Key（選填）", type="password", placeholder="sk-...")
    st.caption("用於「🤖 AI 智能綜合分析」，金鑰只會從您的本機直接呼叫 OpenAI，不會被儲存或上傳到任何伺服器。")
    openai_model = st.selectbox("AI 模型", ["gpt-4o-mini", "gpt-4o"],
                                 format_func=lambda x: {"gpt-4o-mini": "gpt-4o-mini（快速／經濟）",
                                                         "gpt-4o": "gpt-4o（進階／較貴）"}[x])

    st.divider()
    st.markdown("**評分維度各25分**")
    st.caption("📈 趨勢分析（轉折波）")
    st.caption("🕯️ K線型態分析")
    st.caption("📊 均線系統分析")
    st.caption("📦 成交量分析")
    st.markdown("**進場判斷**")
    st.caption("🟢 80+ 積極做多")
    st.caption("🔵 65-79 可考慮進場")
    st.caption("🟡 50-64 觀望")
    st.caption("🔴 <50 不適合進場")


# ────────────────────────────────────────────────────────────────
# 批次分析主流程
# ────────────────────────────────────────────────────────────────

def run_batch_analysis():
    if not api_token:
        st.error("請輸入 FinMind API Token")
        return

    stocks, seen = [], set()
    for tok in st.session_state["stock_text_input"].replace("，", ",").replace("、", ",").split():
        for s in tok.split(","):
            s = s.strip()
            if s and s not in seen:
                seen.add(s)
                stocks.append(s)
    if not stocks:
        st.error("請輸入至少一個股票代號")
        return

    status = st.empty()
    progress_bar = st.progress(0.0)
    log_box = st.container()

    status.text("🔌 測試 API 連線中...")
    try:
        api_fetch(f"{FINMIND_BASE}?dataset=TaiwanStockInfo&token={api_token}")
    except Exception as e:
        st.error(f"❌ 無法連線至 FinMind API\n\n錯誤：{e}\n\n請確認 Token 是否正確、網路連線是否正常。")
        return

    name_map = fetch_stock_name_map(api_token)

    # 盤中即時股價快照：整批股票一次（或分批）呼叫，避免每檔各打一次API
    realtime_map, realtime_err = {}, None
    if use_realtime:
        status.text("🔴 抓取盤中即時股價快照中...")
        realtime_map, realtime_err = fetch_realtime_snapshots(api_token, stocks)
        with log_box:
            if realtime_err:
                st.caption(f"⚠️ 即時快照抓取失敗（已略過，僅用歷史資料分析）：{realtime_err}　※此功能限FinMind sponsor會員使用")
            else:
                st.caption(f"🔴 已取得 {len(realtime_map)} 檔即時快照")

    batch_results = []
    total = len(stocks)
    for i, sid in enumerate(stocks):
        status.text(f"📡 分析 {sid}… ({i + 1}/{total})")
        try:
            rows = fetch_price_data(sid, api_token, days)
            raw_data = sorted(
                [{"date": d["date"], "open": float(d["open"]), "high": float(d["max"]),
                  "low": float(d["min"]), "close": float(d["close"]), "volume": float(d["Trading_Volume"])}
                 for d in rows],
                key=lambda x: x["date"],
            )
            rt_injected = False
            if use_realtime and sid in realtime_map:
                raw_data, rt_injected = merge_realtime_snapshot(raw_data, realtime_map[sid])
            name = name_map.get(sid, sid)
            pe_range = None
            if show_pe:
                try:
                    pe_range = fetch_pe_range(sid, api_token, 3)
                except Exception:
                    pass  # PER抓不到就顯示無資料，不影響其他分析
            rev_range = None
            if show_rev:
                try:
                    rev_range = fetch_revenue_yoy_mom(sid, api_token)
                except Exception:
                    pass  # 營收抓不到就顯示無資料，不影響其他分析
            price_yoy_range = None
            if show_pxyoy:
                try:
                    price_yoy_range = fetch_monthly_avg_price_yoy(sid, api_token, rev_range)
                except Exception:
                    pass  # 均價YoY抓不到就顯示無資料，不影響其他分析
            inst_range = None
            if show_inst:
                try:
                    inst_range = fetch_institutional_monthly(sid, api_token)
                except Exception:
                    pass  # 法人買賣超抓不到就顯示無資料，不影響其他分析
            data = enrich(raw_data)
            dmi = calc_dmi(data, 14)
            dm = score_dmi(dmi)
            pb = check_pullback_buy(data)
            pt = detect_patterns(data, pb)
            total_score = dm["score"]
            batch_results.append({"stockId": sid, "name": name, "data": data, "dm": dm,
                                   "pb": pb, "pt": pt, "total": total_score,
                                   "realtime": rt_injected, "peRange": pe_range,
                                   "revRange": rev_range, "priceYoYRange": price_yoy_range,
                                   "instRange": inst_range})
            with log_box:
                st.caption(f"✅ {sid} {name}　得分:{total_score}" + ("　🔴即時" if rt_injected else ""))
        except Exception as ex:
            with log_box:
                st.caption(f"❌ {sid} 失敗：{ex}")
        progress_bar.progress((i + 1) / total)
        if i < total - 1:
            time.sleep(0.6)

    progress_bar.empty()
    status.empty()

    if not batch_results:
        st.error("所有股票均無法取得資料")
        return

    batch_results.sort(key=lambda r: r["total"], reverse=True)
    st.session_state.batch_results = batch_results
    st.session_state.selected_stock_idx = 0


if run_clicked or st.session_state.pop("_run_after_top100", False):
    run_batch_analysis()


# ────────────────────────────────────────────────────────────────
# 結果顯示
# ────────────────────────────────────────────────────────────────

if "batch_results" not in st.session_state:
    st.info("📈 請在左側輸入 FinMind API Token 與股票代號，點擊「批次分析」即可開始。")
else:
    batch_results = st.session_state.batch_results

    st.markdown("### 📋 批次分析摘要")

    fcol1, fcol2, fcol3, fcol4 = st.columns([1, 1, 1, 1.4])
    with fcol1:
        pb_filter = st.selectbox("進場條件", ["全部", "✅ 符合進場", "❌ 不符合"], key="pb_filter")
    with fcol2:
        score_filter = st.selectbox("評分", ["全部", "80+", "65-79", "50-64", "<50"], key="score_filter")
    with fcol3:
        pt_filter = st.selectbox("型態確認", ["全部", "✅ 已突破", "🔥 剛突破", "🕒 成形中", "－ 無"], key="pt_filter")
    with fcol4:
        kw = st.text_input("搜尋代號/名稱", key="kw_filter", placeholder="輸入代號或名稱關鍵字")

    def build_summary_row(i, r):
        score_lbl = "積極做多" if r["total"] >= 80 else "可考慮進場" if r["total"] >= 65 else "觀望" if r["total"] >= 50 else "不建議進場"
        pb_txt = "符合進場" if r["pb"]["allPass"] else f"{r['pb']['requiredPassed']}/{r['pb']['requiredTotal']} 通過"
        pt_names = [("🔥" if x["justBroke"] else "") + x["name"] for x in r["pt"]["results"]
                    if (x["breakout"] if r["pt"]["anyBreakout"] else x["formed"])]
        pt_txt = "、".join(pt_names) if pt_names else "無"
        pt_icon = "🔥" if r["pt"]["anyJustBroke"] else ("✅" if r["pt"]["anyBreakout"] else ("🕒" if r["pt"]["anyFormed"] else "－"))
        last = r["data"][-1]
        prev = r["data"][-2] if len(r["data"]) >= 2 else last
        chgp = (last["close"] - prev["close"]) / prev["close"] * 100 if prev["close"] else 0
        pe = r.get("peRange")
        pe_txt = "無資料"
        if pe:
            pe_txt = f"{pe['current']:.1f}（{pe['min']:.1f}~{pe['max']:.1f}）"

        def fmt_month_pct(entries, field):
            if not entries:
                return "無資料"
            latest = entries[-1]
            latest_v = latest.get(field)
            main = f"{latest['month']}月 {'+' if latest_v is not None and latest_v>=0 else ''}{latest_v:.1f}%" if latest_v is not None else f"{latest['month']}月 N/A"
            prior = list(reversed(entries[:-1]))
            sub_parts = []
            for m in prior:
                v = m.get(field)
                sub_parts.append(f"{m['month']}月 " + (f"{'+' if v>=0 else ''}{v:.1f}%" if v is not None else "N/A"))
            return main + ("　" + "　".join(sub_parts) if sub_parts else "")

        rev = r.get("revRange")
        yoy_txt = fmt_month_pct(rev, "yoy")
        mom_txt = fmt_month_pct(rev, "mom")
        px_range = r.get("priceYoYRange")
        pxyoy_txt = fmt_month_pct(px_range, "yoy")

        # YoY乖離度＝營收YoY − 均價YoY，3個月都算（不是只算最新月）。
        # 用(year,month)配對，不用陣列位置對應，避免兩邊月份萬一沒對齊時算錯。
        # 主要顯示改成「近3月加總乖離度」（3個月各自的乖離度加總）——單一月份容易
        # 受單月雜訊干擾，加總後比較能看出持續性的乖離趨勢。3個月各自的乖離度數字
        # 還是保留顯示（在加總值下面）。正值大＝營收成長比股價快；負值大＝股價漲幅
        # 超前營收成長。不用多打API，純算既有資料。
        div_txt = "需同時勾營收與均價YoY"
        div_total = None  # 供排序使用
        if rev and px_range:
            px_by_key = {(p["year"], p["month"]): p for p in px_range}
            div_list = []
            for rv_e in rev:
                px_e = px_by_key.get((rv_e["year"], rv_e["month"]))
                d = (rv_e["yoy"] - px_e["yoy"]) if (rv_e.get("yoy") is not None and px_e and px_e.get("yoy") is not None) else None
                div_list.append({"month": rv_e["month"], "div": d})

            def fmt_div(d):
                return f"{'+' if d>=0 else ''}{d:.1f}pp" if d is not None else "N/A"

            valid_divs = [d["div"] for d in div_list if d["div"] is not None]
            if valid_divs:
                div_total = sum(valid_divs)
                lbl = ("💚 營收優於股價" if div_total > 45
                       else "⚠️ 股價超前營收" if div_total < -45 else "大致同步")
                main = f"近3月合計 {fmt_div(div_total)}　{lbl}"
            else:
                main = "當月資料不足"
            sub_parts = [f"{d['month']}月 {fmt_div(d['div'])}" for d in reversed(div_list)]
            div_txt = main + ("　" + "　".join(sub_parts) if sub_parts else "")

        dm = r["dm"]
        # 布林通道股價位置：跟個股詳細面板用同一套算法（不用多打API，last["bbU"]/
        # last["bbL"] 在enrich()時就已經算好了）。0%=貼著下軌，100%=貼著上軌，
        # 50%=通道中央。
        bb_pos_txt = "無資料"
        if last.get("bbU") is not None and last.get("bbL") is not None:
            bb_width = last["bbU"] - last["bbL"]
            bb_pos = ((last["close"] - last["bbL"]) / bb_width * 100) if bb_width else 50
            bb_width_pct = (bb_width / last["close"] * 100) if last["close"] else 0
            bb_pos_txt = f"{bb_pos:.0f}%（寬度{bb_width_pct:.1f}%）"

        # ── MACD狀態 ＋ 布林通道×MACD 情境判斷（強勢突破盤／跌深反彈盤）──
        # 這是課程常見的「布林通道找位置、MACD做確認」的組合判斷法：
        #   強勢突破盤＝通道開口放大＋股價貼近上軌（布林找爆發）
        #             ＋MACD在零軸上且紅柱持續增長（MACD確認不是假突破）
        #   跌深反彈盤＝股價貼近下軌（布林找相對低點）
        #             ＋MACD低檔背離＋（近期）黃金交叉（MACD確認下跌動能衰竭）
        # 低檔背離是真的背離偵測（用跟型態辨識、圖表轉折波同一套 build_zigzag
        # 找最近兩個轉折低點比較），不是代理指標。
        macd_state_txt = "無資料"
        combo_tag = ""
        if (last.get("macd") is not None and last.get("macdSig") is not None
                and last.get("macdHist") is not None and prev.get("macdHist") is not None):
            above_zero = last["macd"] > 0
            hist_growing = last["macdHist"] > prev["macdHist"]
            just_golden_cross = (prev.get("macd") is not None and prev.get("macdSig") is not None
                                  and prev["macd"] <= prev["macdSig"] and last["macd"] > last["macdSig"])
            # 近3天內的黃金交叉（不是只看今天這一根，容許背離成形後晚1-2天才交叉確認）
            golden_cross_recent = False
            rdata = r["data"]
            for gci in range(max(1, len(rdata) - 3), len(rdata)):
                gc_cur, gc_prev = rdata[gci], rdata[gci - 1]
                if (gc_cur.get("macd") is not None and gc_cur.get("macdSig") is not None
                        and gc_prev.get("macd") is not None and gc_prev.get("macdSig") is not None
                        and gc_prev["macd"] <= gc_prev["macdSig"] and gc_cur["macd"] > gc_cur["macdSig"]):
                    golden_cross_recent = True
                    break

            if above_zero and last["macdHist"] > 0:
                macd_state_txt = "零軸上・紅柱增長" if hist_growing else "零軸上・紅柱縮短"
            elif not above_zero and last["macdHist"] < 0:
                macd_state_txt = "零軸下・綠柱縮短" if hist_growing else "零軸下・綠柱增長"
            else:
                macd_state_txt = "交叉轉換中"
            if just_golden_cross:
                macd_state_txt += "　⚡剛黃金交叉"

            # 通道開口是否放大：跟5天前的通道寬度比較
            width_expanding = False
            if last.get("bbU") is not None and last.get("bbL") is not None:
                width_now_pct = (last["bbU"] - last["bbL"]) / last["close"] * 100 if last["close"] else 0
                ref_idx = len(rdata) - 6
                ref_bar = rdata[ref_idx] if ref_idx >= 0 else None
                if ref_bar and ref_bar.get("bbU") is not None and ref_bar.get("bbL") is not None and ref_bar["close"]:
                    width_ref_pct = (ref_bar["bbU"] - ref_bar["bbL"]) / ref_bar["close"] * 100
                    width_expanding = width_now_pct > width_ref_pct
            bb_pos_for_combo = None
            if last.get("bbU") is not None and last.get("bbL") is not None and (last["bbU"] - last["bbL"]):
                bb_pos_for_combo = (last["close"] - last["bbL"]) / (last["bbU"] - last["bbL"]) * 100

            is_breakout = (bb_pos_for_combo is not None and bb_pos_for_combo >= 80 and width_expanding
                           and above_zero and last["macdHist"] > 0 and hist_growing)

            # MACD低檔背離：股價創更低但DIF沒有跟著創新低，只看最近60根K棒內的低點
            divergence_detected = False
            zz_for_div = build_zigzag(rdata)
            zz_lows = [p for p in zz_for_div if p["type"] == "L"]
            if len(zz_lows) >= 2:
                recent_low, prior_low = zz_lows[-1], zz_lows[-2]
                within_lookback = recent_low["idx"] >= len(rdata) - 1 - 60
                macd_at_recent = rdata[recent_low["idx"]].get("macd") if recent_low["idx"] < len(rdata) else None
                macd_at_prior = rdata[prior_low["idx"]].get("macd") if prior_low["idx"] < len(rdata) else None
                if within_lookback and macd_at_recent is not None and macd_at_prior is not None:
                    price_lower_low = recent_low["price"] < prior_low["price"]
                    macd_higher_low = macd_at_recent > macd_at_prior
                    divergence_detected = price_lower_low and macd_higher_low

            is_pullback_rebound = (bb_pos_for_combo is not None and bb_pos_for_combo <= 20
                                    and divergence_detected and golden_cross_recent)

            if is_breakout:
                combo_tag = "　🚀 強勢突破盤"
            elif is_pullback_rebound:
                combo_tag = "　🎯 跌深反彈盤"

        row = {
            "_idx": i, "股票": f"{r['stockId']} {r['name']}" + (" 🔴即時" if r.get("realtime") else ""), "多方力道": r["total"], "評等": score_lbl,
            "+DI": round(dm["plusDI"], 1) if dm["plusDI"] is not None else None,
            "-DI": round(dm["minusDI"], 1) if dm["minusDI"] is not None else None,
            "ADX": round(dm["adx"], 1) if dm["adx"] is not None else None,
            "ADXR": round(dm["adxr"], 1) if dm["adxr"] is not None else None,
            "布林通道位置": bb_pos_txt,
            "MACD狀態": macd_state_txt + combo_tag,
            "漲跌%": round(chgp, 2), "收盤": round(last["close"], 1),
        }
        if show_pe:
            row["近3年P/E區間"] = pe_txt
        if show_rev:
            row["近3月營收YoY"] = yoy_txt
        if show_pxyoy:
            row["近3月均價YoY"] = pxyoy_txt
        if show_rev and show_pxyoy:
            row["近3月YoY乖離度"] = div_txt
        if show_rev:
            row["近3月營收MoM"] = mom_txt

        # 三大法人買賣超（近3月，逐月加總）：主要顯示改成「近3月合計」，跟其他
        # 「近3月」欄位維持一致的呈現方式，下面小字保留3個月各自數字。
        if show_inst:
            inst_range = r.get("instRange")
            if inst_range:
                inst_total = sum(m["net"] for m in inst_range)
                sub_parts = [f"{m['month']}月{'+' if m['net']>=0 else ''}{round(m['net']):,}" for m in reversed(inst_range)]
                row["三大法人買賣超(近3月)"] = f"近3月合計{'+' if inst_total>=0 else ''}{round(inst_total):,}　" + "　".join(sub_parts)
            else:
                row["三大法人買賣超(近3月)"] = "無資料"

        row["型態確認"] = f"{pt_icon} {pt_txt}"
        return row

    def row_passes_filter(r):
        ok_pb = pb_filter == "全部" or (pb_filter == "✅ 符合進場" and r["pb"]["allPass"]) or (pb_filter == "❌ 不符合" and not r["pb"]["allPass"])
        ok_score = (score_filter == "全部"
                    or (score_filter == "80+" and r["total"] >= 80)
                    or (score_filter == "65-79" and 65 <= r["total"] < 80)
                    or (score_filter == "50-64" and 50 <= r["total"] < 65)
                    or (score_filter == "<50" and r["total"] < 50))
        ok_kw = (not kw) or (kw in r["stockId"]) or (kw in r["name"])
        ok_pt = (pt_filter == "全部"
                 or (pt_filter == "✅ 已突破" and r["pt"]["anyBreakout"])
                 or (pt_filter == "🔥 剛突破" and r["pt"]["anyJustBroke"])
                 or (pt_filter == "🕒 成形中" and r["pt"]["anyFormed"] and not r["pt"]["anyBreakout"])
                 or (pt_filter == "－ 無" and not r["pt"]["anyFormed"]))
        return ok_pb and ok_score and ok_kw and ok_pt

    filtered_indices = [i for i, r in enumerate(batch_results) if row_passes_filter(r)]
    st.caption(f"顯示 {len(filtered_indices)} / {len(batch_results)} 檔")

    summary_rows = [build_summary_row(i, batch_results[i]) for i in filtered_indices]
    df_summary = pd.DataFrame(summary_rows)

    if df_summary.empty:
        st.info("沒有符合篩選條件的股票。")
    else:
        st.dataframe(
            df_summary.drop(columns=["_idx"]), use_container_width=True, hide_index=True, height=360,
        )

        # 匯出Excel：跟畫面上的表格一致——目前的篩選、排序後的資料都會反映在匯出結果裡。
        # 需要 openpyxl 套件（pandas寫.xlsx用的引擎），環境裡沒裝的話這裡會噴錯，
        # 跑 `pip install openpyxl` 補上即可。
        try:
            excel_buf = io.BytesIO()
            df_summary.drop(columns=["_idx"]).to_excel(excel_buf, index=False, engine="openpyxl", sheet_name="批次分析摘要")
            st.download_button(
                "📥 匯出 Excel",
                data=excel_buf.getvalue(),
                file_name=f"批次分析摘要_{datetime.today().strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except ImportError:
            st.caption("⚠️ 匯出Excel需要 openpyxl 套件，請先執行 `pip install openpyxl`")

        # 選擇要看詳細分析的股票（取代原本 HTML 版的分頁 tab）
        options = [f"{r['stockId']} {r['name']}（{r['total']}分）" for r in batch_results]
        default_idx = st.session_state.get("selected_stock_idx", 0)
        selected_label = st.selectbox("選擇個股查看詳細分析", options, index=default_idx, key="stock_selector")
        sel_idx = options.index(selected_label)
        st.session_state.selected_stock_idx = sel_idx
        r = batch_results[sel_idx]

        # ── 個股詳細分析 ──
        total = r["total"]
        if total >= 80:
            vc, vt = "#00c864", "強力買進訊號"
        elif total >= 65:
            vc, vt = "#2196f3", "可考慮進場"
        elif total >= 50:
            vc, vt = "#f0a500", "觀望為主"
        else:
            vc, vt = "#ff3c3c", "不建議進場"

        last = r["data"][-1]
        prev = r["data"][-2] if len(r["data"]) >= 2 else last
        chg = last["close"] - prev["close"]
        chgp = (chg / prev["close"] * 100) if prev["close"] else 0

        st.markdown(f"## 🏷️ {r['stockId']} {r['name']}")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("最新收盤", f"{last['close']:.1f}", f"{chg:+.2f} ({chgp:+.2f}%)")
        m2.metric("當日成交量", f"{last['volume']:,.0f}")
        vr = (last["volume"] / last["vm20"]) if last["vm20"] else 1
        m3.metric("量比 vs MA20", f"{vr:.2f}x", "放量" if vr > 1.2 else ("縮量" if vr < 0.8 else "正常"))
        m4.metric("資料日期", last["date"])

        st.divider()
        st.markdown("### 📊 多方力道評分")
        sc1, sc2 = st.columns([1, 2])
        with sc1:
            st.markdown(
                f"<div style='text-align:center;background:linear-gradient(135deg,#1a1a2e,#16213e);"
                f"border-radius:16px;padding:28px 16px;border:1px solid rgba(255,255,255,.1)'>"
                f"<div style='font-size:64px;font-weight:700;color:{vc}'>{total}</div>"
                f"<div style='font-size:12px;color:#888;margin-top:6px'>多方力道 / 100</div>"
                f"<div style='margin-top:12px;display:inline-block;padding:7px 16px;border-radius:8px;"
                f"background:{vc}22;color:{vc};font-weight:700'>{vt}</div></div>",
                unsafe_allow_html=True,
            )
        with sc2:
            dm = r["dm"]
            dims = [("🧭 方向性 (+DI vs -DI)", dm["diPts"], 50, "+DI相對-DI的優勢程度"),
                    ("💪 趨勢強度 (ADX)", dm["adxPts"], 30, "ADX值，越高趨勢越明確"),
                    ("🚀 趨勢動能 (ADX vs ADXR)", dm["adxrPts"], 20, "ADX>ADXR代表趨勢正在轉強")]
            for t, score_val, max_val, d in dims:
                pct = score_val / max_val * 100 if max_val else 0
                col = "#00c864" if pct >= 70 else "#f0a500" if pct >= 40 else "#ff3c3c"
                st.markdown(
                    f"<div style='background:rgba(255,255,255,.03);border:1px solid rgba(255,255,255,.08);"
                    f"border-radius:10px;padding:10px 14px;margin-bottom:6px'>"
                    f"<div style='display:flex;justify-content:space-between;align-items:center'>"
                    f"<div><div style='font-size:12px;color:#888'>{t}</div><div style='font-size:11px;color:#666'>{d}</div></div>"
                    f"<div style='font-size:22px;font-weight:700;color:{col}'>{score_val:.1f}<span style='font-size:11px;color:#555'>/{max_val}</span></div>"
                    f"</div><div style='background:rgba(255,255,255,.05);border-radius:4px;height:5px;margin-top:6px'>"
                    f"<div style='background:{col};border-radius:4px;height:5px;width:{pct}%'></div></div></div>",
                    unsafe_allow_html=True,
                )
            if dm["plusDI"] is not None:
                st.caption(f"+DI={dm['plusDI']:.1f}　-DI={dm['minusDI']:.1f}　"
                           f"ADX={dm['adx']:.1f}　ADXR={dm['adxr']:.1f}" if dm["adx"] is not None and dm["adxr"] is not None
                           else f"+DI={dm['plusDI']:.1f}　-DI={dm['minusDI']:.1f}")

        st.divider()
        st.markdown("### 🔔 DMI訊號")
        dm_sigs = r["dm"]["sigs"]
        if dm_sigs:
            for label, kind in dm_sigs:
                color = "#00c864" if kind == "bull" else "#ff5555" if kind == "bear" else "#aaa"
                st.markdown(f"<span style='display:inline-block;padding:2px 8px;border-radius:12px;"
                            f"font-size:11px;color:{color};border:1px solid {color}55;margin:2px'>{label}</span>",
                            unsafe_allow_html=True)
        else:
            st.caption("無明顯訊號")

        st.divider()
        st.markdown("### 💡 操作建議")
        dm = r["dm"]
        if total >= 80:
            act, adv = "🟢 積極做多", [f"**{r['name']}** 多方力道評分 {total} 分，DMI顯示多方力道強勁，建議積極做多。"]
            if dm["tdir"] == "多頭":
                adv.append("+DI大於-DI且趨勢轉強，順勢操作，逢低分批佈局。")
            ma20v = last["ma20"] or last["close"]
            bb_up = last["bbU"] or last["close"] * 1.1
            sl = f"停損設於 **{ma20v * 0.97:.1f}** 元（20MA下方3%）"
            tgt = f"目標參考 **{last['close'] * ((bb_up - last['close']) / last['close'] + 1):.1f}** 元（布林上軌）"
        elif total >= 65:
            act, adv = "🔵 可考慮進場", [f"**{r['name']}** 評分 {total} 分，DMI偏多，可考慮分批進場。", "建議等待回測均線後再進場，降低風險。"]
            ma20v = last["ma20"] or last["close"]
            sl = f"停損建議 **{ma20v * 0.98:.1f}** 元（20MA下方2%）"
            tgt = f"短線目標 **{last['close'] * 1.08:.1f}** 元（+8%）"
        elif total >= 50:
            act, adv = "🟡 觀望為主", [f"**{r['name']}** 評分 {total} 分，DMI訊號混雜或趨勢不明確，建議觀望。", "等待ADX轉強或+DI/-DI方向明確後再行動。"]
            sl, tgt = "暫不建議進場", "等待更佳時機"
        else:
            act, adv = "🔴 不適合進場", [f"**{r['name']}** 評分 {total} 分，DMI偏空，不建議進場。"]
            if dm["tdir"] == "空頭":
                adv.append("目前-DI大於+DI，空方力道較強，切忌逆勢做多，等待趨勢反轉。")
            else:
                adv.append("DMI指標偏弱或資料不足，應持現金等待機會。")
            sl, tgt = "持倉者建議設停損出場", "等待多頭訊號出現"

        hints = [s[0] for s in dm["sigs"] if s[1] == "bull"][:5]
        st.markdown(f"**{act}**")
        for line in adv:
            st.markdown(line)
        st.markdown(f"🛑 **停損：**{sl}")
        st.markdown(f"🎯 **目標：**{tgt}")
        if hints:
            st.markdown("✅ **多頭訊號：**" + " · ".join(hints))

        if st.button("🤖 AI 智能綜合分析", key=f"ai_btn_{r['stockId']}"):
            if not openai_key:
                st.warning("請先在左側輸入 OpenAI API Key，才能使用 AI 智能綜合分析。")
            else:
                with st.spinner("正在請 AI 綜合研判技術面數據，請稍候…"):
                    ai_text = run_ai_analysis(r, openai_key, openai_model)
                st.markdown(ai_text)

        st.divider()
        st.markdown("### 🎯 回後買上漲 · 進場條件核對")
        if r["pb"]["allPass"]:
            st.success(f"✅ 符合進場條件（必要條件全部通過）" + ("　+成交量增加分" if r["pb"]["bonusPassed"] else ""))
        else:
            st.error(f"❌ 不符合進場條件（必要條件 {r['pb']['requiredPassed']}/{r['pb']['requiredTotal']} 通過）")
        for cond in r["pb"]["results"]:
            icon = "✅" if cond["pass"] else ("❌" if cond["required"] else "—")
            tag = "" if cond["required"] else " `加分`"
            detail = f"　*(​{cond['detail']})*" if cond.get("detail") else ""
            st.markdown(f"{icon} {cond['label']}{tag}{detail}")

        st.divider()
        st.markdown("### 🔍 型態確認（15種進場型態）")
        if r["pt"]["anyJustBroke"]:
            st.warning("🔥 偵測到剛突破買點（較前一交易日新增）")
        elif r["pt"]["anyBreakout"]:
            st.success("✅ 偵測到型態突破買點")
        elif r["pt"]["anyFormed"]:
            st.info("🕒 型態成形中，尚未突破確認")
        else:
            st.caption("－ 目前未偵測到符合的進場型態")
        for p in r["pt"]["results"]:
            icon = "🔥" if p["justBroke"] else ("✅" if p["breakout"] else ("🕒" if p["formed"] else "－"))
            st.markdown(f"{icon} **{p['name']}**　_{p['desc']}_")
            if p.get("detail"):
                st.caption(p["detail"])

        st.divider()
        st.markdown("### 📐 指標對照")
        ic1, ic2, ic3 = st.columns(3)
        with ic1:
            st.write("**RSI 指標**")
            rv = last["rsi"]
            if rv is not None:
                st.metric("RSI", f"{rv:.1f}")
                st.caption("⚠️ 超買區（>70），注意回調" if rv > 70 else ("💚 超賣區（<30），留意反彈" if rv < 30 else "位於正常區間（30-70）"))
        with ic2:
            st.write("**KD 指標**")
            if last["kdK"] is not None and last["kdD"] is not None:
                kd_color = "🟢" if last["kdK"] > last["kdD"] else "🔴"
                st.markdown(f"K=**{last['kdK']:.1f}**　D=**{last['kdD']:.1f}** {kd_color}")
                st.caption("K>D 偏多" if last["kdK"] > last["kdD"] else "K<D 偏空")
            else:
                st.caption("資料不足")
        with ic3:
            st.write("**均線對照**")
            ma_rows = []
            for label, key in [("MA5", "ma5"), ("MA10", "ma10"), ("MA20", "ma20"), ("MA60", "ma60")]:
                if last[key] is not None:
                    diff = (last["close"] - last[key]) / last[key] * 100
                    ma_rows.append({"均線": label, "數值": round(last[key], 2),
                                     "股價偏離": f"{diff:+.2f}% ({'上方' if diff > 0 else '下方'})"})
            st.dataframe(pd.DataFrame(ma_rows), hide_index=True, use_container_width=True)

        ic4, ic5 = st.columns(2)
        with ic4:
            st.write("**MACD 指標**")
            if last.get("macd") is not None and last.get("macdSig") is not None:
                macd_bull = last["macd"] > last["macdSig"]
                cross = ""
                if prev.get("macd") is not None and prev.get("macdSig") is not None:
                    if prev["macd"] <= prev["macdSig"] and last["macd"] > last["macdSig"]:
                        cross = "　⚡ 黃金交叉"
                    elif prev["macd"] >= prev["macdSig"] and last["macd"] < last["macdSig"]:
                        cross = "　⚡ 死亡交叉"
                st.markdown(f"DIF=**{last['macd']:.2f}**　Signal=**{last['macdSig']:.2f}**　柱狀=**{last['macdHist']:.2f}**")
                st.caption(("DIF在Signal上方，偏多" if macd_bull else "DIF在Signal下方，偏空") + cross)
            else:
                st.caption("資料不足")
        with ic5:
            st.write("**布林通道**")
            if last.get("bbU") is not None and last.get("bbL") is not None:
                bb_width = last["bbU"] - last["bbL"]
                bb_pos = ((last["close"] - last["bbL"]) / bb_width * 100) if bb_width else 50
                bb_width_pct = (bb_width / last["close"] * 100) if last["close"] else 0
                st.markdown(f"上軌=**{last['bbU']:.2f}**　下軌=**{last['bbL']:.2f}**")
                bb_lbl = "⚠️ 貼近上軌，注意過熱回檔" if bb_pos > 80 else ("💚 貼近下軌，留意反彈" if bb_pos < 20 else "位於通道中段")
                if bb_width_pct < 8:
                    bb_lbl += "　🔸通道收窄，留意變盤"
                st.caption(f"股價位置：{bb_pos:.0f}%（通道寬度 {bb_width_pct:.1f}%）　{bb_lbl}")
            else:
                st.caption("資料不足")

        st.divider()
        st.markdown("### 📉 技術分析圖表")
        st.plotly_chart(draw_chart(r["data"], f"{r['stockId']} {r['name']}", r["pt"]), use_container_width=True)

        with st.expander("📋 原始資料（最近20筆）"):
            raw_rows = []
            for d in list(reversed(r["data"]))[:20]:
                raw_rows.append({
                    "日期": d["date"], "開盤": d["open"], "最高": d["high"], "最低": d["low"], "收盤": d["close"],
                    "成交量": d["volume"],
                    "MA5": round(d["ma5"], 2) if d["ma5"] is not None else None,
                    "MA20": round(d["ma20"], 2) if d["ma20"] is not None else None,
                    "MA60": round(d["ma60"], 2) if d["ma60"] is not None else None,
                    "RSI": round(d["rsi"], 2) if d["rsi"] is not None else None,
                    "KD-K": round(d["kdK"], 2) if d["kdK"] is not None else None,
                    "KD-D": round(d["kdD"], 2) if d["kdD"] is not None else None,
                    "MACD": round(d["macd"], 2) if d.get("macd") is not None else None,
                    "Signal": round(d["macdSig"], 2) if d.get("macdSig") is not None else None,
                    "BB上軌": round(d["bbU"], 2) if d.get("bbU") is not None else None,
                    "BB下軌": round(d["bbL"], 2) if d.get("bbL") is not None else None,
                })
            st.dataframe(pd.DataFrame(raw_rows), hide_index=True, use_container_width=True)
