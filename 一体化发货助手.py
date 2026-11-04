#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
周二周五一体化发货助手（V4.12：模板顺序匹配 + 精简终端 + 云仓08补运费清单）
运行：python 一体化发货助手.py
示例：python 一体化发货助手.py --orders "E:/meb/02周二周五发货/01一键搞定发货信息/2026/10月/10.27" --once

只读原始Excel，不操作中通网站/新运营中心。始终先检查生成文件再导入。
环境：pip install pandas openpyxl
"""
from __future__ import annotations
import base64
from io import BytesIO
import argparse
from collections import Counter, defaultdict
import unicodedata
from datetime import date, datetime
import hashlib
import json
from pathlib import Path
import re
import sys
import traceback
import shutil
import pandas as pd

# 无客户信息的中通门店16列空白模板，内置于单文件程序，防止单独从桌面运行时丢失模板。
# 若程序旁/分类根目录有正式中通模板，仍优先使用外部模板。
BUILTIN_ZTO_TEMPLATE_B64 = (
    'UEsDBBQAAAAIACZkX10YOGRp1QAAACoBAAAPAAAAeGwvd29ya2Jvb2sueG1sjc8xTgMxEAXQq1jTs96NQkCr9aZJk5YbOPZs1ortWXmc4AtwDAoK'
    'JDougMRpknNECYi0dKP/pa833bIELw6Y2FFU0FQ1CIyGrItbBfs83D3Csu9K+0xptyHaiRJ85LYoGHOeWinZjBg0VzRhLMEPlILOXFHaSp4Sassj'
    'Yg5ezup6IYN2ES5715T/LhF1QAXHz6/jy/vp4+30+g3i2qytggZEap1V8LSxjZ3fN3bxYOwcZxZ+Pek/HhoGZ3BFZh8w5h9QQq+zo8ijmxiE7Dt5'
    'w8nb3/0ZUEsDBBQAAAAIACZkX118nH/FWAEAAPoDAAANAAAAeGwvc3R5bGVzLnhtbK1Ty2rDMBD8FaF7IzulpYQogRYMveSSHnpV7HUiWElGkoPd'
    'ry+SFdvJoS960ux4d/bp9bZTSM5gnTSa03yRUQK6NJXUR05bX9890e1m3a2c7xH2JwBPOoXarTpOT943K8ZceQIl3MI0oDuFtbFKeLcw9shcY0FU'
    'LoQpZMsse2RKSE2DYm20d6Q0rfacLkcqJvsgZ4Gc5jklLBBaKBioF2FRehN5NkVc3sPg/1uBBFyUkoi3dUnE8DbCe7C6kIgk4be+AU610TAqJudv'
    'g45W9Pny4TougVjJwdgK7E0tAzl0NWE2ege+BMR9WNh7PUbnMbqriW5VofxrxWlGSej6AiVigoNUMgb1ueQlxUz9/q/qXT2l+b1APhMQTYP9rlUH'
    'sEU8wfg5soXRc0tiOI1kPUexaH9ZQv6DHv6lhDTs2Zzj3G8WOfIknDanu5APr+c5X5uL5vQXbz4BUEsDBBQAAAAIACZkX136XAFZAwMAANoNAAAT'
    'AAAAeGwvdGhlbWUvdGhlbWUxLnhtbL1X23KbMBT8FUbvDTdz84RkEsduH9Jpp8kPyCBAjRAeSY6dv+8gbgKM4zR27AdLYs/ZReewwte3+5xor4hx'
    'XNAQmFcG0BCNihjTNARbkXzzwe3NNZyLDOVIozBHIVhkUHz//Qy0fU4on8MQZEJs5rrOowzlkF8VG0T3OUkKlkPBrwqW6jGDO0zTnOiWYbh6DjEF'
    'bd4lQTmigpcLEWFP0QGy8lr8YpY//I0vCNNeIQnBDtO42D2jvQAagVwsCAuBIT9A02+u9TaKiIlgJXAlP01gHRG/WDKQpes20lha/szsGCSCiDFw'
    '6ZffLqNEwChCtJajgk3HNXyrASuoangge+CZ9iBAYbDHDIF7b836ARJVDWfjG10FywenHyBR1dAZBdwZ1n1g9wMkqhq6o4DZ8s6zlv0AicoIpi9j'
    'uOv5vtvAW0xSkB8H8YHrGt5Dg+9gutJqVQIqeo33K0lwhGTf5fBvwVYFFbLKUGCqibcNSmBUNigkeM2w9ojTTEgeOEfwHUDEjwL0AWeO6bsCjlAf'
    'IW3pOgZd3Qy5NbmYfCQTTMiTeCPokUtxvCA4XmFC5ERGtaXYZAvCGsIeMGWwG/M6Vcq1TcFDYIDJXNJBMBXVmus1Tz2ck23+s4jrpjdbO4BzDkV3'
    'wXAUn2gZ5CzlqoYSd7IOz57Q0dENddgn6pB3crIQ3/ywkOCoEF0pD8FUg+Up4cxqu+URJCguC1Yn6JX1LCUOZlN3ZH12a08oMc9gjJq8xpSSqWbr'
    'uvAMRVakeP5hJUEwIaTcqksUWR/bAaH9mbYr+b3m7v7LLDaMiwfIswonL7XnK1VoAsP5Ahqr3JnL0ejDPURJgiIxsdJNH7mosxy8/Fl0OSm2ArGn'
    'LN5pa7Jlf2AcAsczHQNoMeaiKYAWY9a1z/j9oluHZJPB2sl7D22Fl+OWUxEr5Qyl9+e14nW6Ostx9X7UwLWm7NabfhIvcD4Gyrmk+Efgf9RTK6s8'
    '97Gp6lDlTRqtPSHPvpDRdl35dYY6bNnSY5vXMTkb/IFqVm7+AVBLAwQUAAAACAAmZF9dDR656GUAAABzAAAAFAAAAHhsL3NoYXJlZFN0cmluZ3Mu'
    'eG1sBcFRCsMgDADQq0j+Z9w+xpDankXatAomFpMNj7/3lm1ycz8aWrskePoAjmTvR5UrwdfOxwe2dZlR1dzkJhpngmJ2R0TdC3FW32+Sye3sg7Op'
    '7+NCvQflQwuRccNXCG/kXAUcrn9QSwMEFAAAAAgAJmRfXdL0MTFRAgAAyQgAABgAAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWyd1k1v0zAYAOC/'
    'YvlOk/R71dIJGGN8DCYunEPjJtGSuHLcjyOIj5ZDVaS1Egd6GUhw6iohWmkt4s8kZT31LyC7SSqkZNFy8uvGz/u60Wu3+wcdywQtRBwD2zKUMiIE'
    'yK5h1bA1GTZp/U4ZHlT3O5U2JmeOjhAFHcu0nUpHhjqljYogODUdWYqTwQ1kdyyzjomlUCeDiSY4DYIUlTPLFLKiWBQsxbAhS8g/PeKLTwlQUV1p'
    'mvQFbh8jQ9OpDKUCBAJbWMOm44/AMtgmIbCUDh/bhkp1GUp5CHRDVZEtQxGCWtOh2HrpP9ul2fKsz7PpeM7nuXQ87/N8yLO34QWfF9JVL/q8mI6X'
    'fF4KeS57C172eTld9T2f7+2qi7fgkhj0jZiuvhQ2XsrOk4LWY0GqBEHzsSBVgqD9WJAqQdCALEiVIGhBFgQnILGJhN01wO+NQ4UqbEJwGxC2iNdg'
    '4V0JAocfbipDhxL+pFW9nlx4/ZE3mLNUrW3C0NyLNt7lO3cxc6+uvO/n3qd+lLyfIP8Of11fjqPkYVLNL1Nv/DpKPoiWq+EsabdHCTJ+tw+Tasbu'
    '9jhJ9kfu78jdPop5Q+dvYr7e45hSo+m6O4gCT6LButtfdwebZe9M2yw/RsGn0dD9M3YX882y571/GyNPYuTi62o4u1E+i3kdvam7+HyjfB4jJxer'
    '3tz71l39/BHFTmPY6IM3mUUwgR9HPv53RhuKhk4Uohm2A0xUpzIUMyUIyPZnnscUN3hUgOAVphRbwUxHiooIm+UgqGNMw8n2Ugj/lVT/AVBLAwQU'
    'AAAAAAAmZF9d60FwaygBAAAoAQAACwAAAF9yZWxzLy5yZWxz77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48UmVsYXRp'
    'b25zaGlwcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9yZWxhdGlvbnNoaXBzIj48UmVsYXRp'
    'b25zaGlwIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvb2Zm'
    'aWNlRG9jdW1lbnQiIFRhcmdldD0iL3hsL3dvcmtib29rLnhtbCIgSWQ9IlJiYmQxNmE0YTZkZTc0ZGI1IiAvPjwvUmVsYXRpb25zaGlwcz5QSwME'
    'FAAAAAgAJmRfXfYPUdIRAQAA8gIAABoAAAB4bC9fcmVscy93b3JrYm9vay54bWwucmVsc7WSTU7DMBBGr2J5TxwHJySoaTds2JZewLHHsVX/RLYL'
    '6dlYcCSugCgIJYgFm25m8Y309ObTvL++bXazs+gZYjLB95gWJUbgRZDGjz0+ZXXT4t12swfLswk+aTMlNDvrU491ztM9IUlocDwVYQI/O6tCdDyn'
    'IsSRTFwc+QikKsuGxCUDr5nocJ7gP8SglBHwEMTJgc9/gEnKZwsJowOPI+Qek9l+Z8XsLEaPssf7ktadaKq6lYyz9rbDiFxNKGtwsPa5RF+TLqwa'
    '2ZWKNbRUTDEp1DWtkuYR5FOOxo+/21quFnodrzm0qhoGWrGWt9fUewnxmDRAXqv9xJ8HAORle4OkktVUNndCMqjkRY+sPnf7AVBLAwQUAAAACAAm'
    'ZF9djYLZqRYBAABTAwAAEwAAAFtDb250ZW50X1R5cGVzXS54bWytk0FOwzAQRa8SeYtqpywQQkm7ALaABBewnEli1R5bnmlIz8aCI3EFVAdFgJAi'
    '1G48m/F7/y/m4+292o7eFQMksgFrsZalKABNaCx2tdhzu7oW2031cohAxegdUi165nijFJkevCYZIuDoXRuS10wypE5FbXa6A3VZllfKBGRAXvGR'
    'ITbVHbR677i4Hxlw0o7eieJ22juqaqFjdNZotgHVgM0vySq0rTXQBLP3gCwpJtAN9QDsncxTem3xIoPVn84Ejv4n/WolE7i8Q72NNCseB0jJNlA8'
    '6cQP2kMt1OgU8cEByTM3zNAlNffgYXrXJwfImMWyvU7QPHOy2J2983f2UpDXkHb5I6k8Tu//M8zMn4OofCKbT1BLAQIUAxQAAAAIACZkX10YOGRp'
    '1QAAACoBAAAPAAAAAAAAAAAAAACkgQAAAAB4bC93b3JrYm9vay54bWxQSwECFAMUAAAACAAmZF9dfJx/xVgBAAD6AwAADQAAAAAAAAAAAAAApIEC'
    'AQAAeGwvc3R5bGVzLnhtbFBLAQIUAxQAAAAIACZkX136XAFZAwMAANoNAAATAAAAAAAAAAAAAACkgYUCAAB4bC90aGVtZS90aGVtZTEueG1sUEsB'
    'AhQDFAAAAAgAJmRfXQ0euehlAAAAcwAAABQAAAAAAAAAAAAAAKSBuQUAAHhsL3NoYXJlZFN0cmluZ3MueG1sUEsBAhQDFAAAAAgAJmRfXdL0MTFR'
    'AgAAyQgAABgAAAAAAAAAAAAAAKSBUAYAAHhsL3dvcmtzaGVldHMvc2hlZXQxLnhtbFBLAQIUAxQAAAAAACZkX13rQXBrKAEAACgBAAALAAAAAAAA'
    'AAAAAACkgdcIAABfcmVscy8ucmVsc1BLAQIUAxQAAAAIACZkX132D1HSEQEAAPICAAAaAAAAAAAAAAAAAACkgSgKAAB4bC9fcmVscy93b3JrYm9v'
    'ay54bWwucmVsc1BLAQIUAxQAAAAIACZkX12NgtmpFgEAAFMDAAATAAAAAAAAAAAAAACkgXELAABbQ29udGVudF9UeXBlc10ueG1sUEsFBgAAAAAI'
    'AAgAAwIAALgMAAAAAA=='
)

CATEGORIES = {
    '01': '01一键搞定发货信息',
    '03': '03中通打单导入',
    '05': '05中通打单导出',
    '07': '07新运管导入',
    '08': '08中通云仓导入',
    '09': '09中通云仓导出',
}
# 旧脚本 exact-name 清单；新品务必先人工核查渠道。
CLOUD_NAMES = ['【会员权益】【术后礼包】美呗专享',
 '【会员权益】【术后礼包】美呗会员专享',
 '【会员权益】【术后礼包】会员专享小美盒',
 '【6月内购】美呗日记医用重组I型人源化胶原蛋白敷料5片/盒【械】',
 '【调研礼品】美呗日记医用重组I型人源化胶原蛋白敷料5片/盒【械】',
 '【积分兑换】美呗日记医用重组I型人源化胶原蛋白敷料5片/盒【械】(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)',
 '【积分兑换】美呗严选-重组胶原蛋白修复贴（三种包装随机发货）(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)',
 '【超级会员】专属生日礼',
 '【钻石呗】专属生日礼',
 '【黑金呗】专属生日礼',
 '【至尊呗】专属生日礼',
 '【金小呗】专属生日礼',
 '【玉小呗】专属生日礼']
# 旧脚本商品统计重命名表：与历史文件对比时沿用。
SUMMARY_NAMES = {'【会员权益】【术后礼包】美呗专享': '01【会员权益】【术后礼包】美呗专享',
 '【会员权益】【术后礼包】美呗会员专享': '02【会员权益】【术后礼包】美呗会员专享',
 '【会员权益】【术后礼包】会员专享小美盒': '03【会员权益】【术后礼包】会员专享小美盒',
 '【会员权益】【术后礼包】美呗会员专享小美盒': '04【会员权益】【术后礼包】美呗会员专享小美盒',
 '【会员权益】【术后礼包】美呗会员专享修护小美盒': '05【会员权益】【术后礼包】美呗会员专享修护小美盒',
 '【超级会员】专属生日礼': '06【超级会员】专属生日礼',
 '【至尊呗】专属生日礼': '06【至尊呗】专属生日礼',
 '【黑金呗】专属生日礼': '06【黑金呗】专属生日礼',
 '【钻石呗】专属生日礼': '06【钻石呗】专属生日礼',
 '【铂金呗】专属生日礼': '07【铂金呗】专属生日礼',
 '【积分兑换】独特艾琳浪漫奔赴香氛礼盒(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)': '07【积分兑换】独特艾琳浪漫奔赴香氛礼盒',
 '【金小呗】专属生日礼': '08【金小呗】专属生日礼',
 '【积分兑换】英国Cathkidston身体乳护手霜礼盒(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)': '08【积分兑换】英国Cathkidston身体乳护手霜礼盒',
 '【玉小呗】专属生日礼': '09【玉小呗】专属生日礼',
 '【积分兑换】乐普莱柏温和净透洁颜蜜1.3ml*30支/盒(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)': '09【积分兑换】乐普莱柏温和净透洁颜蜜',
 '【银小呗】专属生日礼': '10【银小呗】专属生日礼',
 '【积分兑换】papoo盈润香氛身体乳-桃野仙踪（花果木质调）(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)': '10【积分兑换】papoo盈润香氛身体乳-桃野仙踪',
 '【积分兑换】papoo盈润香氛身体乳-茶园邂逅（自然茶香调）(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)': '10【积分兑换】papoo盈润香氛身体乳-茶园邂逅',
 '【积分兑换】papoo氨基酸香氛沐浴露(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)': '11【积分兑换】papoo氨基酸香氛沐浴露',
 '【积分兑换】papoo氨基酸净透洁面乳-初夏沁梨(100g+50g)(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)': '12【积分兑换】papoo氨基酸净透洁面乳',
 '【积分兑换】papoo闰唇膏-滋润款(moist日用-线下款包装）(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)': '13【积分兑换】papoo闰唇膏-日用',
 '【积分兑换】papoo闰唇膏-润护款(nourish夜用-线下款装)(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)': '14【积分兑换】papoo闰唇膏-夜用',
 '【积分兑换】薇诺娜柔润保湿霜80g(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)': '15【积分兑换】薇诺娜柔润保湿霜',
 '【积分兑换】美呗日记医用妇科护垫（二类器械灭菌姨妈巾替卫生巾）((新疆、西藏、内蒙、海南、阿坝、甘孜 需+10))': '16【积分兑换】美呗日记医用妇科护垫',
 '【积分兑换】美呗严选-重组胶原蛋白修复贴（三种包装随机发货）(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)': '17【积分兑换】重组胶原蛋白修复贴（三种包装随机发货）',
 '【积分兑换】美呗日记医用重组I型人源化胶原蛋白敷料5片/盒【械】(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)': '17【积分兑换】美呗日记胶原蛋白敷料（三种包装随机发货）',
 '【调研礼品】美呗日记医用重组I型人源化胶原蛋白敷料5片/盒【械】': '17【调研礼品】美呗日记医用重组I型人源化胶原蛋白敷料5片/盒【械】',
 '【6月内购】美呗日记医用重组I型人源化胶原蛋白敷料5片/盒【械】': '17【6月内购】美呗日记医用重组I型人源化胶原蛋白敷料5片/盒【械】',
 '【积分兑换】papoo石斛控油去屑香氛洗发露-蜜语苍澜(果香西普调)(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)': '18【积分兑换】papoo石斛香氛洗发露-蜜语苍澜',
 '【积分兑换】papoo石斛控油蓬松香氛洗发露-空林橙语(苍兰果香调)(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)': '18【积分兑换】papoo石斛香氛洗发露-空林橙语',
 '【积分兑换】绽妍重组贻贝粘蛋白速修焕能套装(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)': '19【积分兑换】绽妍重组贻贝粘蛋白速修焕能套装',
 '【积分兑换】颐莲御龄紧颜臻享礼盒(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)': '20【积分兑换】颐莲御龄紧颜臻享礼盒',
 '【积分兑换】泊诗蔻花园系列温莎花园秘境礼盒（50g*3）(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)': '21【积分兑换】泊诗蔻花园系列温莎花园秘境礼盒（50g*3）',
 '【积分兑换】泊诗蔻护手霜中样礼盒莫奈花园礼盒（12g*6）(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)': '22【积分兑换】泊诗蔻护手霜中样礼盒莫奈花园礼盒（12g*6）',
 '【积分兑换】润培云霞庄园系列护手霜套盒（12g*12）(新疆、西藏、内蒙、海南、阿坝、甘孜 需+10)': '23【积分兑换】润培云霞庄园系列护手霜套盒（12g*12）'}
SENSITIVE_AREA = ('新疆', '西藏', '海南', '内蒙', '阿坝', '甘孜')
HMT_AREA = ('香港', '澳门', '澳門', '港澳', '台湾', '臺灣')
ORDER_ALIASES = {'订单号': ['订单号','订单编号','商家订单号','外部订单号'],
                 '姓名': ['收货人姓名','收件人姓名','收件人','姓名'],
                 '电话': ['收货人手机号','收件人手机号','收件人手机','手机','电话'],
                 '地址': ['收货人地址','收件人地址','收货地址','地址']}
WAYBILL_ALIASES = {'订单编号':['订单编号','订单号','商家订单号','外部订单号','客户订单号'],
                   '物流单号':['运单编号','物流单号','快递单号','运单号'],
                   '姓名':['收件人','收件人姓名','收货人姓名','姓名'],
                   '电话':['收件人手机号','收件人手机','收件人电话','收货人手机号','收货人电话','手机','电话']}
OUTPUT_COLS = ['物流公司名称', '物流单号', '订单编号']
ZHONGTONG_COLS = ['收件人姓名','手机','地址','收件人单位','品名']
# 兼容截图里的新版中通门店导入模板（收件人电话/收件人地址）与旧版脚本（手机/地址）。
TEMPLATE_COL_ALIASES = {
    '姓名': ('收件人姓名','收货人姓名','收件人','姓名'),
    '电话': ('收件人电话','收件人手机','收件人手机号','收货人手机号','手机','电话'),
    '地址': ('收件人地址','收件人详细地址','收货人地址','地址'),
    '收件人单位': ('收件人单位','收货人单位','单位'),
    '品名': ('品名','物品名称','商品名称'),
    '数量': ('数量','件数','物品数量'),
    '订单编号': ('订单号','订单编号','商家订单号'),
    '商家备注': ('商家备注','卖家备注','商户备注'),
}
AUDIT_COLS = ['收件人姓名','手机','地址','收件人单位','品名','是否港澳台','是否特殊','是否需要补运费','订单号']



# V4.12：默认只在终端显示业务摘要、异常、待办；完整原始输出写到本地日志。
# --verbose 可恢复完整终端输出（两种模式执行同一业务逻辑）。
VERBOSE_OUTPUT = False
DETAIL_LOG_FILE = None


def _console_important(s):
    line = str(s).strip()
    if not line:
        return False
    if line.startswith(('==========', '[阶段1]', '[运单核对]')):
        return True
    important_tags = (
        '[暂停]', '[暂缓]', '[等待]', '[需要核对]', '[异常]', '[异常定位]',
        '[姓名差异待核对]', '[未完成]', '[已生成]', '[已生成·待人工复核]',
        '[人工已确认]', '[待人工确认]', '[注意文件日期]', '[提示]',
        '[云仓补运费]', '[云仓运费提醒]', '[云仓核对]', '[云仓旧审核表提醒]',
        '[门店暂停]', '[云仓暂停]', '[门店匹配暂停]', '[云仓匹配暂停]',
        '[门店对账顺序]', '[云仓对账顺序]', '[路径错误]', '[本轮发生错误]',
        '[未切换批次]', '[已切换批次]', '[需补商品]', '[配对变化]',
        '[行序关系变化]',
    )
    if line.startswith(important_tags):
        return True
    summary_phrases = (
        '阶段1完成：', '门店三份Excel齐全：', '云仓审核完成：',
        '：按规则配对', '本次共', '本次运行结束',
        '数据预处理检查无异常', '检测到订单、补录及模板未变化',
        '输入数据或模板与上次运行不一致', '未完成阶段1',
        '注意：', '尚未凑齐本批次', '完整放好后',
        '模板无法自动识别', '已有旧模板', '实际发货顺序已核对：',
        '映射未变', '不覆盖人工修改', '无订单，仍会生成',
        '具体原因请查看：', '[详情日志]',
        '请直接粘贴【01分类',
    )
    return any(x in line for x in summary_phrases)


def _console_trim(s):
    s = str(s).strip()
    # 长绝对路径只会造成终端刷屏：显示文件名；完整路径留在运行日志。
    if ('：' in s or ': ' in s) and ('.xlsx' in s or '.xls' in s):
        prefix, sep, suffix = s.partition('：')
        if sep and (suffix.startswith('/') or re.match(r'^[A-Za-z]:[\\/]', suffix)):
            s = prefix + '：' + Path(suffix.replace('\\', '/')).name
    if len(s) > 230:
        return s[:225] + ' …（全文见运行日志）'
    return s


def msg(s):
    message = str(s)
    if DETAIL_LOG_FILE is not None:
        try:
            with DETAIL_LOG_FILE.open('a', encoding='utf-8') as stream:
                stream.write(message + '\n')
        except OSError:
            # 日志磁盘异常不应让对账逻辑半途失败；终端仍可显示关键事项。
            pass
    if VERBOSE_OUTPUT or _console_important(message):
        print(message if VERBOSE_OUTPUT else _console_trim(message), flush=True)

def clean(v):
    if pd.isna(v): return ''
    s = str(v).strip().replace('\u3000','')
    s = re.sub(r'^([0-9]+)\.0$',r'\1',s)
    return s

def key(v): return re.sub(r'\s+','',clean(v))
def unsafe_id(s): return bool(re.search(r'\d+(?:\.\d+)?[eE][+-]?\d+',s))
def excel_safe(v):
    s=clean(v)
    return "'"+s if s.startswith(('=','+','-','@')) else s

def canonical_cols(df):
    df=df.copy()
    df.columns=[str(x).strip().replace('\u3000','') for x in df.columns]
    return df

def aliases(df, lookup):
    rename={}
    for dest, opts in lookup.items():
        if dest in df.columns: continue
        for label in opts:
            if label in df.columns:
                rename[label]=dest
                break
    return df.rename(columns=rename)

def load_excel(p, kind):
    # 前十行寻找真正的列标题，兼容原小程序导出的多行说明。
    raw=pd.read_excel(p, header=None, dtype=str, keep_default_na=False)
    if raw.empty: raise ValueError(f'{p.name} 无数据')
    expect = {'订单号','商品名称','UserId'} if kind=='list' else {'订单号','商品名称'}
    header=None
    for i in range(min(12,len(raw))):
        names=[clean(x).replace('\ufeff','') for x in raw.iloc[i].tolist()]
        if len(set(names)&expect)>=2 and '订单号' in names:
            header=i; break
    if header is None: raise ValueError(f'{p.name} 找不到表头：{sorted(expect)}')
    cols=[clean(x).replace('\ufeff','') or f'空列_{j}' for j,x in enumerate(raw.iloc[header].tolist())]
    data=raw.iloc[header+1:].copy()
    data.columns=cols
    data=data.loc[:,~pd.Index(data.columns).duplicated()]
    data=data[data['订单号'].map(clean).ne('')].copy()
    data['__来源文件']=p.name
    return data.reset_index(drop=True)

def unit_count(spec):
    # 沿用原逻辑：5件套按一套，不将克重/毫升看作包装份数。
    s=clean(spec)
    if re.search(r'\d+件套',s):return 1
    m=re.search(r'(\d+)(?=[包盒袋套件瓶](?!(?:g|ml|克|毫升)))',s)
    return int(m.group(1)) if m else 1

def self_goods(name,spec):
    # 继承旧 MergeOrderInfo 中的名称映射，避免改动中通打单识别。
    mappings=[
      ('棉柔洗脸巾','【积分兑换】棉柔洗脸巾'),
      ('papoo氨基酸净透洁面乳','【积分兑换】papoo氨基酸洁面乳'),
      ('珂莱诗琥珀乌木沐浴露套盒','【积分兑换】珂莱诗琥珀乌木沐浴露套盒'),
      ('绽妍重组贻贝粘蛋白速修焕能套装','【积分兑换】绽妍重组贻贝粘蛋白速修焕能套装'),
      ('颐莲御龄紧颜臻享礼盒','【积分兑换】颐莲御龄紧颜臻享礼盒'),
      ('泊诗蔻花园系列温莎花园秘境礼盒','【积分兑换】泊诗蔻花园系列温莎花园秘境礼盒（50g*3）'),
      ('泊诗蔻护手霜中样礼盒莫奈花园礼盒','【积分兑换】泊诗蔻护手霜中样礼盒莫奈花园礼盒（12g*6）'),
      ('润培云霞庄园系列护手霜套盒','【积分兑换】润培云霞庄园系列护手霜套盒（12g*12）'),
      ('蜜语苍澜','【积分兑换】papoo香氛洗发露-蜜语苍澜'),
      ('空林橙语','【积分兑换】papoo香氛洗发露-空林橙语'),
      ('papoo闰唇膏-润护款','【积分兑换】papoo闰唇膏-夜用'),
      ('papoo闰唇膏-滋润款','【积分兑换】papoo闰唇膏-日用'),
      ('美呗日记医用妇科护垫','【积分兑换】美呗日记医用妇科护垫'),
      ('英国Cathkidston身体乳护手霜礼盒','【积分兑换】英国Cathkidston身体乳护手霜礼盒'),
      ('美呗严选-重组胶原蛋白修复贴','【积分兑换】重组胶原蛋白修复贴'),
      ('7月内购','7月内购'),
      ('【积分兑换】美呗周边|时尚系列咖啡杯','【积分兑换】美呗周边|时尚系列咖啡杯'),
      ('薇诺娜柔润保湿霜80g','【积分兑换】薇诺娜柔润保湿霜'),
      ('Grace sky 冻干粉组合','【积分兑换】Grace sky 冻干粉组合'),
      ('乐普莱柏温和净透洁颜蜜','【积分兑换】乐普莱柏温和净透洁颜蜜'),
      ('独特艾琳浪漫奔赴香氛礼盒','【积分兑换】独特艾琳浪漫奔赴香氛礼盒'),
      ('帆布包&新年日历组合','【积分兑换】帆布包&新年日历组合'),
      ('【术后礼包】美呗会员专享修护小美盒','5档'),
      ('【术后礼包】美呗会员专享小美盒','4档'),
      ('【术后礼包】会员专享小美盒','3档'),
      ('瑷尔博士益生精研光感身体精华液','【积分兑换】瑷尔博士身体精华液'),
    ]
    if 'papoo氨基酸香氛沐浴露' in name:
        if '晨曦微香' in spec:return '【积分兑换】papoo沐浴露-晨曦微香'
        if '玉龙茶香' in spec:return '【积分兑换】papoo沐浴露-玉龙茶香'
    if 'papoo盈润香氛身体乳' in name:
        if '桃野仙踪' in spec:return '【积分兑换】papoo身体乳-桃野仙踪'
        if '茶园邂逅' in spec:return '【积分兑换】papoo身体乳-茶园邂逅'
    for old,new in mappings:
        if old in name:return new
    return name

def cloud_goods(name):
    if '【6月内购】美呗日记医用重组I型人源化胶原蛋白敷料' in name:return 'TG-001'
    if '【调研礼品】美呗日记医用重组I型人源化胶原蛋白敷料' in name:return 'TG-001'
    if '【积分兑换】美呗日记医用重组I型人源化胶原蛋白敷料' in name:return 'TG-001'
    if '【积分兑换】美呗严选-重组胶原蛋白修复贴' in name:return 'TG-001'
    for kw,code in [
        ('【术后礼包】美呗会员专享','TG-015'),
        ('【术后礼包】美呗专享','TG-014'),
        ('【术后礼包】会员专享小美盒','TG-016'),
        ('【超级会员】专属生日礼','TG-007'),
        ('【钻石呗】专属生日礼','TG-007'),
        ('【黑金呗】专属生日礼','TG-007'),
        ('【至尊呗】专属生日礼','TG-007'),
        ('【金小呗】专属生日礼','TG-008'),
        ('【玉小呗】专属生日礼','TG-010'),
    ]:
        if kw in name:return code
    return ''

def review_record(issues,stage,oid,reason,detail=''):
    issues.append({'阶段':stage,'订单号':oid,'问题':reason,'详情':str(detail)})

def normalize_day(s):
    s=s.strip()
    for fmt in ('%Y-%m-%d','%Y%m%d','%Y.%m.%d','%Y/%m/%d'):
        try:return datetime.strptime(s,fmt).date()
        except ValueError:pass
    raise ValueError('日期格式请写 2026-10-27（历史复现必须写年份）')

def locate_base(root,code):
    # 仅查一级分类文件夹：找到现有分类则沿用其完整名称，不重命名。
    matches=[p for p in root.iterdir() if p.is_dir() and re.match(r'^'+code+r'(?!\d)',p.name)] if root.exists() else []
    if len(matches)>1:raise ValueError(f'{code} 文件夹有多个候选：{matches}；请删除歧义或调整名称')
    p=matches[0] if matches else root/CATEGORIES[code]
    p.mkdir(parents=True,exist_ok=True)
    return p

def paths(root,dt):
    # 先复用已有的 2026/10月/10.27，不创建平行目录。
    result={}
    for code in CATEGORIES:
        base=locate_base(root,code)
        year=base/str(dt.year)
        month=year/(str(dt.month)+'月')
        date_name=f'{dt.month}.{dt.day}'
        leaf=month/date_name
        # 如果已有 10.27 / 10.27 (zero padded) 兼容历史目录
        for m in [month,year/('%02d月'%dt.month)]:
            for n in [f'{dt.month}.{dt.day}',f'{dt.month}.{dt.day:02d}',f'{dt.month:02d}.{dt.day:02d}',f'{dt.month:02d}.{dt.day}']:
                if (m/n).is_dir():leaf=m/n;break
            if leaf.is_dir():break
        leaf.mkdir(parents=True,exist_ok=True)
        result[code]=leaf
    return result

def save(df,p):
    p.parent.mkdir(parents=True,exist_ok=True)
    df.to_excel(p,index=False)
    msg(f'  生成：{p}')

def report(issues, p):
    cols=['阶段','订单号','问题','详情']
    save(pd.DataFrame(issues,columns=cols),p)

def source_files(raw_folder):
    """优先读取 01/日期/订单文件 下的8份原始订单；兼容历史放在日期目录的结构。"""
    raw_folder=Path(raw_folder)
    folder=raw_folder/'订单文件' if (raw_folder/'订单文件').is_dir() else raw_folder
    files=sorted(p for p in folder.glob('订单*.xlsx') if p.is_file() and not p.name.startswith('~$'))
    msg(f'  原始订单读取位置：{folder}（{len(files)}份）')
    return files


def ensure_summary_outside(raw_folder):
    """旧版统计表从合并结果迁移到日期根目录；不覆盖现有文件。"""
    old=Path(raw_folder)/'合并结果'/'商品统计表.xlsx'
    target=Path(raw_folder)/'商品统计表.xlsx'
    if old.is_file() and not target.exists():
        old.replace(target)
        msg('  已把历史商品统计表移至01日期目录：'+str(target))
    elif old.is_file() and target.is_file() and file_digest(old)==file_digest(target):
        # 对已生成的重复旧文件仅做归档，不直接删除。
        archive=Path(raw_folder)/'合并结果'/'旧版统计表备份'
        archive.mkdir(parents=True,exist_ok=True)
        where=archive/'商品统计表.xlsx'
        if where.exists():
            where=archive/f'商品统计表_{datetime.now().strftime("%Y%m%d%H%M%S%f")}.xlsx'
        old.replace(where)
    return target

def input_signature(files):
    # 以文件内容计算指纹，防止源文件更换后误用已经打印的旧模板。
    h=hashlib.sha256()
    for f in sorted(files):
        h.update(f.name.encode('utf-8'))
        with f.open('rb') as stream:
            for chunk in iter(lambda:stream.read(1024*1024),b''):
                h.update(chunk)
    return h.hexdigest()

def invalidate_imports(p,dt):
    # 阶段1失败/变动时，旧的“核对通过”文件必须变为不可用。
    for label in ('门店','门店补运费','云仓'):
        old=p['07']/f'{dt.month}.{dt.day}_新运管导入_{label}_核对通过.xlsx'
        if old.exists():
            archive_invalid_import(old,p)

# V3.3 不再生成「补运费清单_需人工确认」和「补运费人工确认记录」。
# 所有门店订单在一个原始订单审核表中人工确认特殊地区、港澳台和补运费。
def audit_file(p,dt,channel='门店'):
    return p['03' if channel=='门店' else '08']/f'{dt.month}.{dt.day}_{channel}原始订单_人工审核.xlsx'

def audit_marker_file(p,channel='门店'):
    """审核状态存放于01合并结果，不额外污染03/08日期文件夹。旧状态安全迁移。"""
    status_dir=p['01']/'合并结果'/'审核状态'
    status_dir.mkdir(parents=True,exist_ok=True)
    new=status_dir/f'_{channel}审核制单指纹.json'
    previous=p['03' if channel=='门店' else '08']/new.name
    if previous.exists():
        if new.exists():
            # 同一批次旧状态与新状态同时出现且内容不一致，必须人工核查。
            if previous.read_bytes()!=new.read_bytes():
                raise ValueError(f'{channel}旧版审核状态与新版审核状态冲突，请备份后人工核对：{previous} 和 {new}')
            previous.unlink()
        else:
            shutil.move(str(previous),str(new))
    return new

def template_order_map_path(p,channel):
    """映射仅存01合并结果，03/08正式模板不增加任何列或Sheet。"""
    return p['01']/'合并结果'/'审核状态'/f'_{channel}实际发货顺序映射.json'


def _shipping_template_rows(file, channel):
    """读取实际中通模板的可见收件/品名列，不按合并表猜测行序。"""
    from openpyxl import load_workbook
    wb=load_workbook(file,read_only=True,data_only=True)
    try:
        ws=wb.active
        mapping={};header=None
        for rownum,values in enumerate(ws.iter_rows(max_row=min(ws.max_row,15),values_only=True),start=1):
            cells={clean(v).replace(' ',''):j for j,v in enumerate(values) if clean(v)}
            found={k:next((cells[a] for a in names if a in cells),None)
                   for k,names in TEMPLATE_COL_ALIASES.items()}
            if all(found.get(k) is not None for k in ('姓名','电话','地址','品名')):
                mapping=found;header=rownum;break
        if not header:
            raise ValueError(f'{channel}中通发货模板找不到姓名、手机、地址、品名表头')
        output=[]
        fields=('姓名','电话','地址','品名')
        for rownum,values in enumerate(ws.iter_rows(min_row=header+1,values_only=True),start=header+1):
            row={field:clean(values[mapping[field]]) if mapping[field]<len(values) else ''
                 for field in fields}
            if not any(row.values()):continue  # 忽略模板末尾纯空白格式行
            if not all(row.values()):
                raise ValueError(f'{channel}模板第{rownum}行的姓名/手机/地址/品名缺项')
            row['模板Excel行号']=rownum
            output.append(row)
        return output
    finally:
        wb.close()


def _row_content_hash(row):
    """不在JSON中保存客户姓名/手机号/住址原文，只保存识别性摘要与订单号。"""
    fields=[clean(row.get('姓名','')),phone_key(row.get('电话','')),
            clean(row.get('地址','')),clean(row.get('品名',''))]
    return hashlib.sha256(json.dumps(fields,ensure_ascii=False).encode('utf-8')).hexdigest()


def ensure_template_order_map(merged,p,dt,channel,issues):
    """以实际上传模板为准落地行序；发现排序/改名/改品名后不再使用旧映射。"""
    try:
        expected=merged[merged['渠道'].eq(channel)].copy()
        if channel=='门店':
            expected['品名']=expected.apply(
                lambda r:f"【{r['实际数量']}】 * {self_goods(r['商品名称'],r['商品规格'])}",axis=1)
        else:
            expected['仓库货号']=expected['商品名称'].map(cloud_goods)
            expected['品名']=expected.apply(lambda r:f"【{r['实际数量']}】 * {r['仓库货号']}",axis=1)
            expected['是否补运费待确认']=expected['是否偏远地区'] & expected['仓库货号'].eq('TG-001')
        audit,signature=read_verified_audit(expected,p,dt,issues,channel=channel)
        if audit is None:raise ValueError('不能重新验证对应的人工审核表')
        if channel=='门店':
            subset=audit[~audit['是否补运费待确认']]
            template=self_template_paths(p,dt)[0]
        else:
            subset=audit
            template=p['08']/f'{dt.month}.{dt.day}.xlsx'
        if not template.is_file():raise ValueError('找不到实际发货模板：'+template.name)
        current=_shipping_template_rows(template,channel)
        if len(current)!=len(subset):
            raise ValueError(f'实际模板{len(current)}行，审核应发{subset.shape[0]}行；可能手动删除/新增/改序')
        entries=[]
        for ix,((_,order),got) in enumerate(zip(subset.iterrows(),current),start=1):
            # 只认可模板列所示内容与审核订单行序完全一致；非订单信息列可自行编辑。
            h=_row_content_hash(order)
            if h!=_row_content_hash(got):
                raise ValueError(f'第{ix}笔（中通模板Excel第{got["模板Excel行号"]}行）与审核订单内容/行序不一致；'
                                 '禁止按旧顺序配物流。请恢复正确模板或在未发货前重新制单。')
            entries.append({'模板Excel行号':got['模板Excel行号'],
                            '订单号':clean(order['订单号']), '内容摘要':h})
        if len({r['订单号'] for r in entries})!=len(entries):
            raise ValueError('模板顺序映射中出现重复订单号')
        dest=template_order_map_path(p,channel)
        state={'version':1,'渠道':channel,'日期':str(dt),'审核签名':signature,
               '模板文件名':template.name,'记录':entries}
        if dest.exists():
            old=json.loads(dest.read_text('utf-8'))
            if old.get('version')!=1 or old.get('渠道')!=channel or old.get('日期')!=str(dt):
                raise ValueError('已有顺序映射所属版本/渠道/日期不一致')
            if old.get('记录')!=entries:
                # F/G/H人工修改导致模板安全重新生成时，才允许更新映射。
                # 单纯改模板的顺序，审核签名不变，必须拒绝而不是重新学习错误顺序。
                marker=audit_marker_file(p,channel)
                registered=json.loads(marker.read_text('utf-8')) if marker.exists() else {}
                allowed=(old.get('审核签名')!=signature
                         and registered.get('signature')==signature
                         and (registered.get('template_hashes') or {}).get(template.name)==file_digest(template))
                if not allowed:
                    raise ValueError('已保存的模板行号—订单号映射发生变化，疑似手动调换模板行序；禁止自动配对')
                msg(f'  [{channel}映射更新] 审核F/G/H改变且新版模板通过校验，更新原行序映射。')
            else:
                # 订单行序未变时仅更新审核签名；不触碰实际订单关联。
                if old.get('审核签名')==signature and old.get('模板文件名')==template.name:
                    msg(f'  {channel}实际发货顺序已核对：{len(entries)}行，映射未变。')
                    return True
        else:
            # V4.10旧批次安全迁移：只有旧模板和审核表逐行一致才新建映射；不按手机号猜。
            msg(f'  [{channel}映射建立] 按实际发货模板逐行核对{len(entries)}笔订单并保存行号映射。')
        dest.parent.mkdir(parents=True,exist_ok=True)
        tmp=dest.with_name(dest.name+'.tmp')
        tmp.write_text(json.dumps(state,ensure_ascii=False,indent=2),'utf-8')
        tmp.replace(dest)
        return True
    except Exception as exc:
        review_record(issues,channel+'发货顺序','','模板行序映射校验失败',str(exc))
        msg(f'  [{channel}匹配暂停] 模板顺序映射无法确认：{exc}')
        return False


def expected_in_template_order(expected,p,dt,channel):
    """回传对账入口只认可已核验并保存的模板订单行序。"""
    dest=template_order_map_path(p,channel)
    if not dest.is_file():
        raise ValueError(f'{channel}尚无实际发货模板顺序映射，禁止按合并表原行序猜测物流单号')
    info=json.loads(dest.read_text('utf-8'))
    if info.get('渠道')!=channel or info.get('日期')!=str(dt) or info.get('version')!=1:
        raise ValueError('实际发货顺序映射与本批日期或渠道不符')
    records=info.get('记录') or []
    ids=[clean(x['订单号']) for x in records]
    src=expected.copy()
    src['订单号']=src['订单号'].map(clean)
    if (len(ids)!=len(src) or len(set(ids))!=len(ids) or
        len(set(src['订单号']))!=len(src) or set(ids)!=set(src['订单号'])):
        raise ValueError('映射中的订单号集合与本批应发订单不一致；禁止物流配对')
    ordered=src.set_index('订单号',drop=False).loc[ids].reset_index(drop=True)
    ordered['__模板Excel行号']=[int(x['模板Excel行号']) for x in records]
    return ordered


def self_template_paths(p,dt):
    base=f'{dt.month}.{dt.day}_中通门店_'
    return (p['03']/(base+'普通订单.xlsx'),p['03']/(base+'补运费订单_确认后导入.xlsx'))

def audit_table(rows):
    # A:H 是可见人工审核区；订单号存于隐藏 I 列，只用于安全匹配，避免重排后错配。
    return pd.DataFrame([{
        '收件人姓名':clean(r['姓名']),
        '手机':clean(r['电话']),
        '地址':clean(r['地址']),
        '收件人单位':'',
        '品名':clean(r['品名']),
        '是否港澳台':'是' if r['是否港澳台'] else '',
        '是否特殊':'是' if r['是否偏远地区'] else '',
        '是否需要补运费':'补运费10元' if r['是否补运费待确认'] else '',
        '订单号':clean(r['订单号']),
    } for _,r in rows.iterrows()], columns=AUDIT_COLS)


def parse_yes_blank(value, col, oid):
    v=clean(value)
    # 容忍上一版留下的「否」，将其视为空白；其他文本必须核对。
    if v in ('', '否'):return False
    if v=='是':return True
    raise ValueError(f'订单{oid}「{col}」只允许填写「是」或留空，当前是：{v}')


def parse_fee_cell(value, oid):
    # 「是否需要补运费」一格同时表达是否补运费以及金额。
    v=clean(value).replace(' ', '')
    if v in ('','否'):return False,''
    if v=='是':return True,'10'  # 兼容旧版审核表未迁移状态
    m=re.fullmatch(r'(?:需要)?补运费(\d+(?:\.\d{1,2})?)元?',v)
    if not m:raise ValueError(f'订单{oid}「是否需要补运费」应留空或填写「补运费10元」，当前是：{v}')
    from decimal import Decimal
    num=Decimal(m.group(1))
    if num<=0 or num>=100000:raise ValueError(f'订单{oid}补运费金额必须大于0且低于100000元')
    result=format(num.normalize(),'f')
    return True,result


def backup_and_convert_v33_audit(dest,p,channel):
    """已有V3.3审核表自动迁移到V3.4。先在01合并结果保留备份，不额外占03/08的Excel名额。"""
    from openpyxl import load_workbook
    from openpyxl.worksheet.datavalidation import DataValidation
    import shutil
    wb=load_workbook(dest)
    try:
        if channel+'原始订单' not in wb.sheetnames:return
        ws=wb[channel+'原始订单']
        header=[clean(ws.cell(1,c).value) for c in range(1,12)]
        has_extra = '补运费金额（元）' in header or '审核备注' in header
        has_old_no = any(clean(ws.cell(r,c).value)=='否' for r in range(2,ws.max_row+1) for c in (6,7,8))
        # V3.8样式修订：兼容已有的蓝色表头和黄色F/G/H审核列；只清除填充色，不改用户填写内容。
        has_fill = any(cell.fill.patternType not in (None, 'none')
                       for sheet in wb.worksheets for row in sheet.iter_rows() for cell in row)
        if not has_extra and not has_old_no and not has_fill:
            return  # 已经符合「正常留空，无填充色」的新版格式
        backup_dir=p['01']/'合并结果'/'旧版审核表备份'
        backup_dir.mkdir(parents=True,exist_ok=True)
        backup=backup_dir/dest.name
        if not backup.exists():shutil.copy2(dest,backup)
        # 同步旧 F/G/H/J 人工填写值；仅在审核文件上迁移，不修改来源订单。
        for row_idx in range(2,ws.max_row+1):
            for c in (6,7):
                if clean(ws.cell(row_idx,c).value)=='否':ws.cell(row_idx,c).value=None
            h=clean(ws.cell(row_idx,8).value)
            if h in ('','否'):
                ws.cell(row_idx,8).value=None
            elif h=='是':
                raw=(clean(ws.cell(row_idx,10).value) if has_extra else '') or '10'
                try:
                    from decimal import Decimal
                    amount=format(Decimal(raw).normalize(),'f')
                except Exception:
                    raise ValueError(f'旧审核表第{row_idx}行金额无效：{raw}，未修改原文件')
                ws.cell(row_idx,8).value=f'补运费{amount}元'
            elif h.startswith('需要补运费'):
                ws.cell(row_idx,8).value=h.replace('需要补运费','补运费',1)
        # 删除多余的「金额」「审核备注」列；第I列订单号仅保留为隐藏匹配键。
        if has_extra:
            ws.delete_cols(10,2)
        ws.column_dimensions['I'].hidden=True
        ws.auto_filter.ref=f'A1:H{ws.max_row}'
        if '审核确认' in wb.sheetnames:
            guide=wb['审核确认']
            guide['A1']='审核说明：查看原始订单的F/G/H列；正常留空，特殊填“是”，补运费在H列写“补运费10元”'
            guide['A4']='1. F港澳台、G特殊地区：正常留空，异常填“是”。'
            guide['A5']='2. H补运费：无需补运费留空；需要补运费填“补运费10元”或对应金额。'
            guide['A6']='3. 隐藏I列为订单号，不要修改；无需再改B2，保存后回程序按回车。'
            guide['A7']='4. 不再需要填写审核状态；请核对F/G/H并保存，然后回程序按回车。'
            guide['B2']=None
        for dv in list(ws.data_validations.dataValidation):
            ws.data_validations.dataValidation.remove(dv)
        valid=DataValidation(type='list',formula1='"是"',allow_blank=True)
        ws.add_data_validation(valid)
        if ws.max_row >= 2:valid.add(f'F2:G{ws.max_row}')
        # 只去除背景填充色；不清空数据、不改变订单号、筛选、字体或手动编辑。
        if has_fill:
            from openpyxl.styles import PatternFill
            no_fill=PatternFill(fill_type=None)
            for sheet in wb.worksheets:
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.fill.patternType not in (None, 'none'):
                            cell.fill=no_fill
        # 在本地文件旁先临时保存，保存完成后原子替换，失败不会损坏旧审核表。
        tmp=dest.with_name('_临时升级_'+dest.name)
        try:
            wb.save(tmp)
            tmp.replace(dest)
        finally:
            if tmp.exists():tmp.unlink()
        msg('  已更新'+channel+'审核表的F/G/H及无填充色格式，备份存放于：'+str(backup))
    finally:
        wb.close()


def ensure_audit_file(rows,p,dt,channel='门店'):
    """只新建缺失的表，旧版审核运行一次无损迁移；不覆盖用户F/G/H人工填写。"""
    dest=audit_file(p,dt,channel)
    if dest.exists():
        backup_and_convert_v33_audit(dest,p,channel)
        return dest
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    from openpyxl.worksheet.datavalidation import DataValidation
    wb=Workbook()
    ws=wb.active;ws.title=channel+'原始订单'
    ws.append(AUDIT_COLS)
    df=audit_table(rows)
    for record in df.itertuples(index=False,name=None):ws.append(list(record))
    ws.freeze_panes='F2';ws.auto_filter.ref=f'A1:H{ws.max_row}'
    widths={'A':16,'B':17,'C':48,'D':16,'E':56,'F':14,'G':12,'H':22,'I':25}
    for k,w in widths.items():ws.column_dimensions[k].width=w
    ws.column_dimensions['I'].hidden=True
    for cell in ws[1]:
        cell.font=Font(bold=True,color='1B365D')
        cell.alignment=Alignment(horizontal='center',vertical='center')
    for row in range(2,ws.max_row+1):
        ws.cell(row,2).number_format='@';ws.cell(row,9).number_format='@'
    option=DataValidation(type='list',formula1='"是"',allow_blank=True)
    ws.add_data_validation(option)
    if ws.max_row>=2:option.add(f'F2:G{ws.max_row}')
    # V3.8 只提供一个审核工作表；无需再设置B2/审核确认。
    dest.parent.mkdir(parents=True,exist_ok=True);wb.save(dest)
    msg('  生成'+channel+'原始订单人工审核表：'+str(dest))
    msg('  请在真正导入中通/发送云仓前检查F/G/H；若有修改，保存后按回车重新核对模板。')
    return dest


def read_verified_audit(expected, p, dt, issues, channel='门店'):
    """审核表存在或新建后读取；校验成功返回(订单行,审核指纹)。"""
    # 第一次执行也直接读取新生成的审核表，并按默认地区识别输出模板。
    # 生成不代表审核完成；用户必须在真正导入中通/发送云仓前自行核对F/G/H。
    dest=audit_file(p,dt,channel)
    existed=dest.is_file()
    dest=ensure_audit_file(expected,p,dt,channel)
    if not existed:
        msg(f'  [需要核对] {channel}新审核表已生成；本轮同时生成待核对的中通模板，发货前必须人工审核。')
    try:
        df=pd.read_excel(dest,sheet_name=channel+'原始订单',dtype=str,keep_default_na=False)
        df=canonical_cols(df)
        absent=[c for c in AUDIT_COLS if c not in df.columns]
        if absent:raise ValueError('审核表缺列：'+','.join(absent))
        for c in AUDIT_COLS:df[c]=df[c].map(clean)
        if df['订单号'].eq('').any() or df['订单号'].duplicated().any():
            raise ValueError('审核表隐藏订单号为空或重复，无法确定审核行归属')
        base=expected.copy()
        if base['订单号'].duplicated().any():raise ValueError('同订单多商品，审核表不能用单订单号定位')
        if set(df['订单号'])!=set(base['订单号']):
            raise ValueError('审核表的订单号与该批原始数据不一致：存在漏行、错行或跨日期混用')
        audit=df.set_index('订单号').loc[base['订单号'].tolist()]
        for audit_col,base_col in [('收件人姓名','姓名'),('手机','电话'),('地址','地址'),('品名','品名')]:
            if audit[audit_col].tolist()!=[clean(x) for x in base[base_col].tolist()]:
                raise ValueError(f'审核表「{audit_col}」与订单原文不一致；只允许调整F/G/H列')
        H=[]; G=[]; fees=[]; amounts=[]
        for oid,f,g,h in zip(base['订单号'],audit['是否港澳台'],audit['是否特殊'],audit['是否需要补运费']):
            H.append(parse_yes_blank(f,'是否港澳台',oid))
            G.append(parse_yes_blank(g,'是否特殊',oid))
            needed,amount=parse_fee_cell(h,oid)
            fees.append(needed);amounts.append(amount)
        base['是否港澳台']=H
        base['是否偏远地区']=G
        base['是否补运费待确认']=fees
        base['补运费金额']=amounts
        # 云仓额外检查地址原文；即使F列被误清空，也不能让港澳地址流入中通。
        if channel=='云仓':
            actual_hm=base['地址'].map(lambda a:any(x in clean(a) for x in ('香港','澳门','澳門','港澳')))
            missed=base.loc[actual_hm & ~base['是否港澳台'],'订单号'].tolist()
            if missed:
                raise ValueError(f'云仓有{len(missed)}条港澳地址，但审核F列未写「是」；请核实地址后单独处理。')
        restricted=base.loc[base['是否港澳台'],'订单号'].tolist()
        if restricted:raise ValueError(f'人工审核标出{len(restricted)}条港澳台订单；请先在原系统单独处理，禁止自动导入中通。')
        # 只对审核内容签名；保存Excel或改变行顺序不需要重新制单。
        content=[{'订单号':clean(r['订单号']),'港澳台':bool(r['是否港澳台']),
                  '特殊':bool(r['是否偏远地区']),'补运费':bool(r['是否补运费待确认']),
                  '金额':clean(r['补运费金额'])} for _,r in base.iterrows()]
        signature=hashlib.sha256(json.dumps(content,ensure_ascii=False,sort_keys=True).encode('utf-8')).hexdigest()
        return base,signature
    except Exception as e:
        review_record(issues,channel+'审核','',channel+'人工审核表校验失败',e)
        msg('  ['+channel+'暂停] '+str(e))
        return None,None

def file_digest(path):
    """保存模板哈希：手工修改过的模板绝不自动覆盖。"""
    h=hashlib.sha256()
    with Path(path).open('rb') as f:
        for chunk in iter(lambda:f.read(1048576),b''):
            h.update(chunk)
    return h.hexdigest()


def make_manifest(signature, dt, templates, counts):
    return {'signature':signature,'date':str(dt),'counts':counts,
            'template_hashes':{x.name:file_digest(x) for x in templates}}


def refresh_after_audit(channel, rows, signature, current_marker, outputs, p, dt, root, issues):
    """只对“本程序生成且未改动、未发货”的模板同步 F/G/H 编辑。

    保存旧模板到01/合并结果/模板变更备份；先在临时目录全部生成成功，再替换正式模板。
    """
    previous_hashes=current_marker.get('template_hashes') or {}
    if not previous_hashes:
        raise ValueError('旧版本制单记录缺少校验哈希；不能自动覆盖。请备份并确认未发货后人工清理旧版模板/指纹，再重新运行。')
    for dest,_ in outputs:
        if dest.exists() and previous_hashes.get(dest.name)!=file_digest(dest):
            raise ValueError(f'{dest.name}被人工修改，禁止覆盖。请检查文件内容并手动备份后再处理。')
    back_folder=p['05' if channel=='门店' else '09']
    returned=find_back_file(back_folder,['人工匹配_'])
    imported=p['07']/f'{dt.month}.{dt.day}_新运管导入_{channel}_核对通过.xlsx'
    if returned or imported.exists():
        raise ValueError(f'{channel}已有物流回传或新运管导入文件，审核表变更不得自动改写可能已发出的模板。')
    # 避免模板在原目录中生成中途残留多份临时Excel。
    staging=p['01']/'合并结果'/'待写入模板'/channel
    staging.mkdir(parents=True,exist_ok=True)
    temp_files=[]
    issue_count=len(issues)
    for dest,subset in outputs:
        tmp=staging/dest.name
        if tmp.exists():tmp.unlink()
        create_template_file(subset,tmp,root,'self' if channel=='门店' else 'cloud',issues)
        if not tmp.exists() or len(issues)!=issue_count:
            raise ValueError('新模板生成失败，原文件未改动。详情请查异常报告。')
        temp_files.append((tmp,dest))
    backup=p['01']/'合并结果'/'模板变更备份'/channel/(datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
    backup.mkdir(parents=True,exist_ok=True)
    for _,dest in temp_files:
        if dest.exists():shutil.copy2(dest,backup/dest.name)
    for temp,dest in temp_files:
        temp.replace(dest)
    msg(f'  {channel}审核表有修改：已更新未发货的模板，原模板已备份于01合并结果。')
    return True


def prepare_self_from_audit(merged,p,dt,root,issues):
    """只在人工确认后生成正式门店模板；再次运行不会覆盖打印前手动调整的文件。"""
    expected=merged[merged['渠道'].eq('门店')].copy()
    expected['品名']=expected.apply(lambda r:f"【{r['实际数量']}】 * {self_goods(r['商品名称'],r['商品规格'])}",axis=1)
    rows,signature=read_verified_audit(expected,p,dt,issues)
    if rows is None:
        stale=p['07']/f'{dt.month}.{dt.day}_新运管导入_门店_核对通过.xlsx'
        if stale.exists():archive_invalid_import(stale,p)
        return False
    normal_file,bonus_file=self_template_paths(p,dt)
    marker=audit_marker_file(p)
    if marker.exists():
        try:
            old=json.loads(marker.read_text(encoding='utf-8'))
            if old.get('signature')!=signature:
                items=[(normal_file,rows[~rows['是否补运费待确认']]),
                       (bonus_file,rows[rows['是否补运费待确认']])]
                refresh_after_audit('门店',rows,signature,old,items,p,dt,root,issues)
                marker.write_text(json.dumps(make_manifest(signature,dt,[normal_file,bonus_file],
                    {'普通订单':len(items[0][1]),'补运费订单':len(items[1][1])}),
                    ensure_ascii=False,indent=2),encoding='utf-8')
                msg('  门店审核更新已同步：审核表、普通单、补运费单共3份。')
                return True
            # 若仅丢失其中一份，由于已有制单指纹且审核内容完全一致，
            # 只补建缺失的那份文件；既有文件原样保留，不改变已发货记录。
            missing = [(normal_file, rows[~rows['是否补运费待确认']]),
                       (bonus_file, rows[rows['是否补运费待确认']])]
            for target, items in missing:
                if not target.exists():
                    before = len(issues)
                    create_template_file(items,target,root,'self',issues)
                    if len(issues) != before or not target.exists():
                        raise ValueError(f'补建缺失的门店Excel失败：{target.name}，请核实实际中通模板及异常报告')
                    msg('  已按当前审核数据补建缺失文件：'+target.name)
            for f in (normal_file,bonus_file):
                normalize_old_template_fill(f,p,'门店')
            if old.get('template_hashes'):
                old['template_hashes']={f.name:file_digest(f) for f in (normal_file,bonus_file)}
                marker.write_text(json.dumps(old,ensure_ascii=False,indent=2),'utf-8')
            msg('  门店日期文件夹应有三份：原始审核、普通订单、补运费订单（无需补运费也生成空表）。')
            msg('  审核结果未变化，保留已存在的模板，不覆盖人工修改。')
            return True
        except Exception as e:
            review_record(issues,'门店审核','','制单保护',e)
            old=p['07']/f'{dt.month}.{dt.day}_新运管导入_门店_核对通过.xlsx'
            if old.exists():archive_invalid_import(old,p)
            msg('  [门店暂停] '+str(e));return False
    if normal_file.exists() or bonus_file.exists():
        # 未找到对应制单指纹时无法证明既有模板属于本批次，禁止覆盖或者
        # 在它旁边自动生成另一个可能与其不一致的正式物流表。
        review_record(issues,'门店审核','','已有门店模板但缺少制单指纹',
                      '请先核查是否已经打单；如为旧版本遗留文件，请备份后移走旧文件再按回车生成完整三表')
        msg('  [门店暂停] 发现旧模板但没有可信的制单记录；为避免重复发货，请备份并移走旧模板后按回车。')
        return False
    normal=rows[~rows['是否补运费待确认']]
    bonus=rows[rows['是否补运费待确认']]
    # 必须两份模板一起生成成功，才把候选文件放进03文件夹。
    # V4.6 在缺少补运费模板时会先写普通单，再将它删除，造成「先生成、后缺少」的误导。
    staging=p['01']/'合并结果'/'待写入模板'/'门店'
    staging.mkdir(parents=True,exist_ok=True)
    temps=[(staging/normal_file.name,normal_file,normal),
           (staging/bonus_file.name,bonus_file,bonus)]
    for tmp,_,_ in temps:
        if tmp.exists():tmp.unlink()
    original_count=len(issues)
    try:
        for tmp,_,subset in temps:
            create_template_file(subset,tmp,root,'self',issues)
            if len(issues)!=original_count or not tmp.exists():
                msg('  [门店暂停] 候选模板未全部生成成功；03原有文件保持不变。')
                return False
        installed=[]
        try:
            for tmp,dest,_ in temps:
                if dest.exists():
                    raise ValueError('门店目标路径在制单过程中已有文件，禁止覆盖：'+str(dest))
                tmp.replace(dest)
                installed.append(dest)
            marker.write_text(json.dumps(make_manifest(signature,dt,[normal_file,bonus_file],
                {'普通订单':len(normal),'补运费订单':len(bonus)}),ensure_ascii=False,indent=2),encoding='utf-8')
        except Exception:
            # 这里的dest都是本轮新创建的，才允许回滚；绝不清理用户原有Excel。
            for dest in installed:
                if dest.exists():dest.unlink()
            raise
        msg(f'  门店三份Excel齐全：审核表1份 + 普通订单{len(normal)}条 + 补运费订单{len(bonus)}条（无补运费也保留空模板）。')
        return True
    finally:
        for tmp,_,_ in temps:
            if tmp.exists():tmp.unlink()

def sync_cloud_fee_checklist(rows,p,dt):
    """08云仓补运费提示表：仅做人工核对，绝不可当作第二份中通发货订单上传。

    只要审核H列明确标记补运费就列出（包含TG-001，也列出人工标记的其他货号）。
    08日期目录正常只有审核表+完整发货表；存在需联系的订单才增加第三份。
    若审核取消补运费，旧清单安全移到01合并结果的备份目录，不留下过期提醒。
    """
    name=f'{dt.month}.{dt.day}_云仓补运费订单_人工确认.xlsx'
    target=p['08']/name
    archive=p['01']/'合并结果'/'云仓补运费清单历史备份'
    legacy=p['01']/'合并结果'/f'{dt.month}.{dt.day}_云仓TG001补运费需联系.xlsx'

    def backup_file(source):
        if not source.exists():return
        archive.mkdir(parents=True,exist_ok=True)
        new=archive/source.name
        if new.exists():
            new=archive/(source.stem+'_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f')+source.suffix)
        shutil.move(str(source),str(new))

    # V4.11 的旧提醒文件不再放在用户平时不看的合并结果目录，原文件仅归档。
    if legacy.exists():backup_file(legacy)

    flagged=rows.loc[rows['是否补运费待确认'].astype(bool)].copy()
    if flagged.empty:
        if target.exists():
            backup_file(target)
            msg('  [云仓补运费] 当前审核H列无补运费订单；08仅保留2份Excel，旧提醒已备份。')
        return

    def is_tg001(val):
        return bool(re.search(r'(?<![A-Za-z0-9])TG-001(?![A-Za-z0-9])',clean(val),re.I))

    frame=pd.DataFrame({
        '订单编号':flagged['订单号'].map(excel_safe),
        '收件人姓名':flagged['姓名'].map(excel_safe),
        '手机':flagged['电话'].map(excel_safe),
        '地址':flagged['地址'].map(excel_safe),
        '品名':flagged['品名'].map(excel_safe),
        '补运费金额（元）':flagged['补运费金额'].map(excel_safe),
        '货号判断':flagged['品名'].map(lambda x:'TG-001' if is_tg001(x) else '其他（人工标记）'),
        '人工确认结果':['待联系云仓确认']*len(flagged),
    })
    compare_cols=['订单编号','收件人姓名','手机','地址','品名','补运费金额（元）','货号判断']
    if target.exists():
        try:
            old=pd.read_excel(target,dtype=str,keep_default_na=False)
            if all(c in old.columns for c in compare_cols):
                left=old[compare_cols].fillna('').astype(str).map(clean).reset_index(drop=True)
                right=frame[compare_cols].fillna('').astype(str).map(clean).reset_index(drop=True)
                if left.equals(right):
                    msg(f'  [云仓补运费] 08已保留{len(flagged)}条补运费人工核对清单；{dt.month}.{dt.day}.xlsx仍保留全部订单。')
                    return
        except Exception:
            # 不覆盖读取失败的表；下方先备份再安全重建。
            pass
        backup_file(target)
        msg('  [云仓补运费] 审核H列与旧清单不同；旧版已备份，重新生成08人工核对清单。')
    target.parent.mkdir(parents=True,exist_ok=True)
    temp=p['01']/'合并结果'/'待写入模板'/'云仓'/name
    temp.parent.mkdir(parents=True,exist_ok=True)
    try:
        save(frame,temp)
        temp.replace(target)
    finally:
        if temp.exists():temp.unlink()
    msg(f'  [云仓补运费] 08新增第3份Excel：{target.name}（{len(flagged)}条，人工确认用，不上传中通）。')
    msg(f'  [云仓补运费] {dt.month}.{dt.day}.xlsx不删任何订单；仅上传此表，不要重复上传人工核对清单。')


def prepare_cloud_from_audit(merged,p,dt,root,issues):
    """08目录基础两份Excel；审核H列标记补运费时增加1份只用于人工核对的清单。"""
    expected=merged[merged['渠道'].eq('云仓')].copy()
    if expected.empty:
        msg('  云仓本批无订单，仍会生成两份只有表头的Excel。')
    expected['仓库货号']=expected['商品名称'].map(cloud_goods)
    unknown=expected.loc[expected['仓库货号'].eq(''),'订单号'].tolist()
    if unknown:
        review_record(issues,'云仓货号','','云仓商品无TG编码',f'{len(unknown)}个订单，禁止生成云仓模板')
        msg('  [云仓暂停] 有商品未匹配TG编码，请先核对云仓货号。')
        return False
    expected['品名']=expected.apply(lambda r:f"【{r['实际数量']}】 * {r['仓库货号']}",axis=1)
    # 云仓默认规则：偏远地区 + 货号TG-001才自动标注补运费。
    # 如审核表已存在，仍以审核表人工修改的H列为准。
    expected['是否补运费待确认']=expected['是否偏远地区'] & expected['仓库货号'].eq('TG-001')
    rows,signature=read_verified_audit(expected,p,dt,issues,channel='云仓')
    if rows is not None:
        tg001=rows['品名'].map(lambda v:bool(re.search(r'(?<![A-Za-z0-9])TG-001(?![A-Za-z0-9])',clean(v),re.I)))
        fee=rows['是否补运费待确认'].astype(bool)
        need=int((tg001 & fee).sum())
        other=int((~tg001 & fee).sum())
        # 用户历史日期目录中可能有上一版本生成的审核表；不覆盖其H列，但提示漏标。
        suspicious=rows.loc[tg001 & rows['是否偏远地区'].astype(bool) & ~fee,'订单号'].tolist()
        if suspicious:
            msg(f'  [云仓旧审核表提醒] {len(suspicious)}条TG-001偏远地区订单H列为空；请核实是否漏填补运费，不会擅自覆盖人工审核。')
        if need:
            msg(f'  [云仓补运费] {need}条【TG-001且审核H列写有补运费】，请与云仓确认实际收费；正式五列发货表保留全部订单。')
        if other:
            msg(f'  [云仓核对] 另有{other}条非TG-001订单填写了补运费；按本次规则不自动按TG-001收费，请人工核实。')
    if rows is None:
        old=p['07']/f'{dt.month}.{dt.day}_新运管导入_云仓_核对通过.xlsx'
        if old.exists():old.unlink()
        return False
    # V4.1：云仓发给中通的文件仅用日期命名，如 10.30.xlsx。
    template=p['08']/f'{dt.month}.{dt.day}.xlsx'
    marker=audit_marker_file(p,'云仓')
    former_names=[p['08']/f'{dt.month}.{dt.day}_云仓订单.xlsx',
                  p['08']/f'{dt.month}.{dt.day}_发中通云仓_中通模板.xlsx']
    existing_former=[f for f in former_names if f.exists()]
    # 已有V4.0/V3.9发货文件时，只有制单指纹能证明其未经修改，才自动更名。
    # 如已手工修改或已有同名文件，暂停以避免覆盖实际已发送的版本。
    if existing_former:
        if template.exists() or len(existing_former)>1:
            review_record(issues,'云仓审核','','发现多个版本云仓发货文件',
                          '请人工核对并备份，保留正确文件后再运行：'+', '.join(f.name for f in existing_former))
            msg('  [云仓暂停] 云仓目录发现同名或多个旧发货文件，未覆盖任何文件。')
            return False
        former=existing_former[0]
        if not marker.exists():
            review_record(issues,'云仓审核','','旧云仓文件缺少制单记录','先备份并核对：'+former.name)
            msg('  [云仓暂停] 旧文件无制单指纹，需人工核实。')
            return False
        try:
            old_json=json.loads(marker.read_text('utf-8'))
            known=(old_json.get('template_hashes') or {}).get(former.name)
            if not known or known!=file_digest(former):
                review_record(issues,'云仓审核','','旧云仓文件已修改/缺少哈希',
                              '为保护人工修改和已发货记录，未自动更名：'+former.name)
                msg('  [云仓暂停] 旧云仓文件可能经过人工修改，需先核查。')
                return False
            former.rename(template)
            old_json['template_hashes'].pop(former.name,None)
            old_json['template_hashes'][template.name]=file_digest(template)
            marker.write_text(json.dumps(old_json,ensure_ascii=False,indent=2),'utf-8')
            msg('  已安全更名云仓文件：'+former.name+' -> '+template.name)
        except Exception as e:
            review_record(issues,'云仓审核','','旧云仓文件安全更名失败',str(e))
            return False
    if marker.exists():
        try:
            old=json.loads(marker.read_text('utf-8'))
            if old.get('signature')!=signature:
                refresh_after_audit('云仓',rows,signature,old,[(template,rows)],p,dt,root,issues)
                marker.write_text(json.dumps(make_manifest(signature,dt,[template],{'云仓订单':len(rows)}),
                    ensure_ascii=False,indent=2),'utf-8')
                sync_cloud_fee_checklist(rows,p,dt)
                msg('  云仓审核更新已同步：发货表完整保留，补运费清单按H列同步。')
                return True
            if not template.exists():
                before=len(issues)
                create_template_file(rows,template,root,'cloud',issues)
                if not template.exists() or len(issues)!=before:raise ValueError('云仓模板丢失后补建失败')
                msg('  已补建缺失的云仓中通模板：'+template.name)
            normalize_old_template_fill(template,p,'云仓')
            if old.get('template_hashes'):
                old['template_hashes']={template.name:file_digest(template)}
                marker.write_text(json.dumps(old,ensure_ascii=False,indent=2),'utf-8')
            sync_cloud_fee_checklist(rows,p,dt)
            msg('  云仓审核内容未变，保留已有中通模板文件。')
            return True
        except Exception as e:
            review_record(issues,'云仓审核','','云仓制单保护',e)
            old=p['07']/f'{dt.month}.{dt.day}_新运管导入_云仓_核对通过.xlsx'
            if old.exists():archive_invalid_import(old,p)
            msg('  [云仓暂停] '+str(e))
            return False
    if template.exists():
        review_record(issues,'云仓审核','','存在旧云仓中通模板','为保护可能已发出的文件，请先备份并移走旧模板')
        msg('  [云仓暂停] 已存在旧云仓模板，不自动覆盖。')
        return False
    before=len(issues)
    create_template_file(rows,template,root,'cloud',issues)
    if len(issues)==before and template.exists():
        marker.write_text(json.dumps(make_manifest(signature,dt,[template],{'云仓订单':len(rows)}),ensure_ascii=False,indent=2),'utf-8')
        sync_cloud_fee_checklist(rows,p,dt)
        msg(f'  云仓审核完成：生成{len(rows)}条中通模板订单。')
        return True
    if template.exists():template.unlink()
    return False

def find_custom_template(root,category):
    """只定位模板原件，不读取/改写模板中的示例行。"""
    if category=='self':
        names = ['模板.xlsx','中通门店模板_空白版.xlsx','中通门店订单导入模板new.xlsx','中通门店订单导入模板.xlsx',
                 '中通门店模板.xlsx','中通门店导入模板.xlsx']
        fallback = ('中通门店','模板')
    else:
        names = ['中通云仓订单导入模板new.xlsx','中通云仓订单导入模板.xlsx',
                 '中通云仓模板.xlsx','云仓模板.xlsx','中通云仓导入模板.xlsx']
        fallback = ('云仓','模板')
    folders = [Path(__file__).resolve().parent,root]
    for folder in folders:
        for name in names:
            candidate=folder/name
            if candidate.is_file(): return candidate
        candidates=sorted(x for x in folder.glob('*.xlsx')
                          if not x.name.startswith('~$') and all(part in x.name for part in fallback))
        if len(candidates)==1:return candidates[0]
    # 云仓按用户截图固定为5列；不借用含多余字段的门店16列模板。
    # 如有独立云仓正式模板，则优先使用该模板。
    return None

def create_template_file(rows, dest, root, which, issues):
    """尽量用真实中通模板的字段名/顺序；无法识别的模板拒绝冒充可导入文件。"""
    # 此处只在程序准备阶段执行一次；后续运单回传运行不会覆盖这份模板。
    rows=rows.copy()
    records=[]
    for _,r in rows.iterrows():
        cloud_tg001=(which=='cloud' and bool(re.search(r'(?<![A-Za-z0-9])TG-001(?![A-Za-z0-9])',clean(r.get('品名','')),re.I)))
        needs_fee=bool(r.get('是否补运费待确认',False)) and (which!='cloud' or cloud_tg001)
        records.append({'订单编号':excel_safe(r['订单号']),'姓名':excel_safe(r['姓名']),'电话':excel_safe(r['电话']),
                        '地址':excel_safe(r['地址']),'收件人单位':'',
                        '品名':excel_safe(r['品名']),'数量':int(r['实际数量']),
                        '商家备注':('需要补运费'+clean(r.get('补运费金额','10') or '10')+'元')
                                    if needs_fee else ''})
    template=find_custom_template(root,which)
    if template or which=='self':
        try:
            # 优先使用外部正式模板；缺失时从程序内部读取无客户信息的16列空白模板。
            from openpyxl import load_workbook
            if template:
                wb=load_workbook(template)
                source_desc=template.name
            else:
                wb=load_workbook(BytesIO(base64.b64decode(BUILTIN_ZTO_TEMPLATE_B64)))
                source_desc='程序内置16列空白备用模板（建议使用上传的模板.xlsx）'
            ws=wb.active
            header_row=None; mappings={}
            for ridx in range(1,min(ws.max_row,15)+1):
                cells={clean(ws.cell(ridx,j).value).replace(' ',''):j for j in range(1,ws.max_column+1)
                       if clean(ws.cell(ridx,j).value)}
                found={field:next((cells[a] for a in opts if a in cells),None)
                       for field,opts in TEMPLATE_COL_ALIASES.items()}
                if all(found[k] is not None for k in ('姓名','电话','地址','品名')):
                    header_row=ridx; mappings={k:j for k,j in found.items() if j is not None}
                    break
            if header_row is None:
                raise ValueError('找不到姓名、电话、地址、品名这四个必备列；请核对模板真实表头')
            if len(set(mappings.values()))!=len(mappings):
                raise ValueError('模板列映射重复，无法确认写入位置')
            if which=='self' and any(r['商家备注'] for r in records) and '商家备注' not in mappings:
                raise ValueError('补运费订单需要在「商家备注」列注明金额，但模板缺少「商家备注」；请使用实际门店模板')
            # 删除模板已有样例数据，避免把样例当作待发货的真实订单。
            # 保留用户模板A:P全部16列、标题、样式和「填写说明」工作表。
            # 原模板可能含客户示例，先清除所有历史示例值，不能混入本次发货。
            from copy import copy
            style_row=[copy(ws.cell(header_row+1,col)._style) for col in range(1,ws.max_column+1)]
            for row_cells in ws.iter_rows(min_row=header_row+1,max_row=ws.max_row):
                for cell in row_cells:
                    cell.value=None
            # 仅删除示例下方的空白数据行，不删除任何列，也不改表头和其他工作表。
            if ws.max_row>header_row+1:
                ws.delete_rows(header_row+2,ws.max_row-header_row-1)
            # 只写入 E姓名 / F电话 / G地址 / I品名四项对应的列。
            # 补运费单沿用之前确认的 P商家备注；普通单其他列保持空白。
            required_fields=('姓名','电话','地址','品名')
            for i,record in enumerate(records,start=header_row+1):
                for col,style in enumerate(style_row,start=1):
                    ws.cell(i,col)._style=copy(style)
                for field in required_fields:
                    cell=ws.cell(i,mappings[field])
                    cell.value=record.get(field,'')
                    if field in ('电话','品名'):
                        cell.number_format='@'
                if which=='self' and record['商家备注'] and '商家备注' in mappings:
                    ws.cell(i,mappings['商家备注']).value=record['商家备注']
            # 完整保留原模板格式，不再清除原有表头样式或删减任何列。
            dest.parent.mkdir(parents=True,exist_ok=True)
            wb.save(dest)
            msg(f'  使用模板：{source_desc}（表头第{header_row}行）')
            msg(f'  生成：{dest}')
            msg('  已保留模板全部16列；仅填写收件人姓名、电话、地址、品名（补运费单另填写商家备注）。')
            return
        except Exception as e:
            review_record(issues,'模板','', '模板不兼容/读取失败',f'{source_desc}: {e}')
            msg(f'  模板无法自动识别：{source_desc}；为了防止错列，没有生成{dest.name}')
            return
    # 未提供云仓专属模板时，严格按截图生成5列，不混入门店长模板。
    if which=='cloud':
        fee_num=sum(bool(r['商家备注']) for r in records)
        if fee_num:
            msg(f'  [云仓运费提醒] 有{fee_num}条满足【H列补运费 + 品名TG-001】。云仓五列模板无商家备注，请单独联系云仓确认收费，不要改动TG货号。')
        msg('  云仓发货订单采用截图中的五列：收件人姓名/手机/地址/收件人单位/品名。')
    else:
        msg('  注意：未找到门店正式模板，暂使用五列表头；补运费单必须使用真实模板。')
    extra_col=['商家备注'] if which=='self' and any(r['商家备注'] for r in records) else []
    standard=pd.DataFrame([{'收件人姓名':r['姓名'], '手机':r['电话'],
                            '地址':r['地址'],'收件人单位':'','品名':r['品名'],
                            '商家备注':r['商家备注']}
                           for r in records],columns=ZHONGTONG_COLS+extra_col)
    # 没有真实门店模板时不允许把带补运费的备用格式假装是可导入文件。
    if which=='self' and any(r['商家备注'] for r in records):
        review_record(issues,'模板','','缺少中通门店正式模板','无法确认「商家备注」应写入的列位置；请把门店模板放到根目录')
        return
    save(standard,dest)

def supplement_recipient(merged, folder, issues):
    """仅补充源订单的空白收件信息；永不静默覆盖真实已有信息。"""
    supplement=folder/'收件信息补录.xlsx'
    missing=merged[["姓名","电话","地址"]].eq('').any(axis=1)
    if missing.any() and not supplement.exists():
        to_fill=merged.loc[missing,['订单号','姓名','电话','地址']].drop_duplicates('订单号').copy()
        to_fill['补录说明']='只填写空白单元格；手机/地址须从原系统核实，不可推测；按订单号对应'
        save(to_fill,supplement)
        msg('  已生成缺失收件信息补录表：'+str(supplement))
    if supplement.exists():
        try:
            add=pd.read_excel(supplement,dtype=str,keep_default_na=False)
            required={'订单号','姓名','电话','地址'}
            if not required.issubset(add.columns):
                raise ValueError('补录表缺列：'+','.join(sorted(required-set(add.columns))))
            for c in required:add[c]=add[c].map(clean)
            if add['订单号'].eq('').any() or add['订单号'].duplicated().any():
                raise ValueError('订单号为空或重复；补录文件必须一订单一行')
            not_found=set(add['订单号'])-set(merged['订单号'])
            if not_found:raise ValueError(f'补录表有{len(not_found)}个未知订单号；不要跨日期混用')
            extra=add.set_index('订单号')
            for col in ['姓名','电话','地址']:
                incoming=merged['订单号'].map(extra[col]).fillna('').map(clean)
                conflict=(merged[col].ne(''))&incoming.ne('')&(merged[col]!=incoming)
                if conflict.any():
                    raise ValueError(f'补录表有{int(conflict.sum())}条「{col}」与原订单冲突，禁止覆盖')
                merged[col]=merged[col].where(merged[col].ne(''),incoming)
        except Exception as e:
            review_record(issues,'收件补录','','补录文件不合法',e)
    return merged, supplement

def stage_prepare(raw_folder,p,dt,root,issues):
    ensure_summary_outside(raw_folder)
    files=source_files(raw_folder)
    lists=[f for f in files if '订单列表' in f.name]
    details=[f for f in files if '订单列表' not in f.name]
    msg(f'[阶段1] 原始文件：总计{len(files)}，小程序{len(lists)}，运营订单{len(details)}')
    if len(files)!=8 or not lists or not details:
        invalidate_imports(p,dt)
        msg('  尚未凑齐本批次8份订单文件，或缺少订单列表/订单类型。请检查：')
        for f in files:msg('    '+f.name)
        msg('  完整放好后按回车重新检查，暂不生成打单/云仓模板。')
        return None
    sig_file=p['01']/'合并结果'/'_源文件指纹.json'
    # 补录表也属于实际制单依据：一旦制单成功，后续修改必须手动新批次确认。
    supplement_file=p['01']/'合并结果'/'收件信息补录.xlsx'
    template_inputs=[find_custom_template(root,x) for x in ('self','cloud')]
    source_sig=input_signature(files+([supplement_file] if supplement_file.exists() else [])+[f for f in template_inputs if f is not None])
    processed=p['01']/'合并结果'/'一键生成py.xlsx'
    if sig_file.exists():
        try:
            previous=json.loads(sig_file.read_text(encoding='utf-8'))
            if source_sig != previous.get('signature'):
                review_record(issues,'订单合并','','订单、补录信息或中通模板有变化','不能覆盖本批次已经打单的数据。确认后可备份旧日期文件夹并删除 _源文件指纹.json 重新生成。')
                invalidate_imports(p,dt)
                msg('  输入数据或模板与上次运行不一致，已停止，保护人工修改过的模板。')
                return None
            if not processed.exists():
                review_record(issues,'订单合并','','历史合并结果丢失','已有源文件指纹但一键生成py.xlsx缺失；请先确认旧批次输出，不自动覆盖')
                invalidate_imports(p,dt)
                return None
            if processed.exists():
                ensure_summary_outside(raw_folder)
                msg('  检测到订单、补录及模板未变化：复用上次合并结果，不覆盖已人工修改的中通/云仓模板。')
                cached=pd.read_excel(processed,dtype=str,keep_default_na=False)
                for c in ('是否补运费待确认','是否偏远地区','是否积分兑换','是否港澳台'):
                    if c in cached:cached[c]=cached[c].map(lambda x:str(x).lower() in ('true','1','是'))
                return cached
        except Exception as e:
            review_record(issues,'订单合并','','无法复用源文件指纹',e)
            return None
    try:
        left=pd.concat([load_excel(f,'list') for f in lists],ignore_index=True)
        right=pd.concat([load_excel(f,'detail') for f in details],ignore_index=True)
    except Exception as e:
        review_record(issues,'订单导出','','原始表无法读取',e)
        return None
    for df,source,need in [(left,'订单列表',['订单号','商品名称','UserId']),
                            (right,'运营订单',['订单号','商品名称','商品规格','数量','收货人姓名','收货人手机号','收货人地址'])]:
        absent=[c for c in need if c not in df]
        if absent:
            review_record(issues,'订单合并','',source+'缺列',','.join(absent))
            return None
        df['订单号']=df['订单号'].map(clean)
        df['商品名称']=df['商品名称'].map(clean)
        df['__商品键']=df['商品名称'].map(key)
    if left['订单号'].map(unsafe_id).any() or right['订单号'].map(unsafe_id).any():
        review_record(issues,'订单合并','','订单号科学计数法','Excel 可能损坏精度；请重新导出原始文件')
        return None
    # 相同订单同商品多行/不同渠道、数据源明细不对应时必须复核，拒绝按订单号多对多合并。
    if left.duplicated(['订单号','__商品键']).any() or right.duplicated(['订单号','__商品键']).any():
        for source,df in [('小程序',left),('运营订单',right)]:
            bad=df[df.duplicated(['订单号','__商品键'],keep=False)]
            for oid in bad['订单号'].unique():review_record(issues,'订单合并',oid,source+'出现重复订单商品','需确认多规格或多次导出重复')
        return None
    lkeys=set(map(tuple,left[['订单号','__商品键']].values.tolist()))
    rkeys=set(map(tuple,right[['订单号','__商品键']].values.tolist()))
    if lkeys != rkeys:
        for oid,sku in sorted(lkeys-rkeys)[:100]:review_record(issues,'订单合并',oid,'小程序商品找不到对应运营订单',sku)
        for oid,sku in sorted(rkeys-lkeys)[:100]:review_record(issues,'订单合并',oid,'运营订单商品找不到对应小程序订单',sku)
        return None
    keep=['订单号','__商品键','商品规格','数量','订单金额','实际支付金额','收货人姓名','收货人手机号','收货人地址']
    for c in keep:
        if c not in right:right[c]=''
    # 关键修复：两类原始文件可能同时包含「商品规格」「数量」「收货人姓名」等列。
    # pandas 默认会改名为 商品规格_x/商品规格_y，导致 merged['商品规格'] KeyError。
    # 规格、数量、金额、收件信息均以运营订单明细(right)为准；小程序(left)用于商品身份/用户信息。
    # 一定在 merge 前消除同名非键列，而不是 merge 后猜测 _x/_y。
    detail_cols=set(keep)-{'订单号','__商品键'}
    overlapping=[c for c in left.columns if c in detail_cols or c in ('姓名','电话','地址')]
    if overlapping:
        msg('  发现两份导出存在同名字段，统一以运营订单明细为准：'+', '.join(overlapping))
    left_safe=left.drop(columns=overlapping,errors='ignore')
    merged=left_safe.merge(right[keep],on=['订单号','__商品键'],how='inner',validate='one_to_one')
    merged=merged.rename(columns={'收货人姓名':'姓名','收货人手机号':'电话','收货人地址':'地址'})
    absent_after_merge=[c for c in ('姓名','电话','地址','数量','商品规格') if c not in merged]
    if absent_after_merge:
        review_record(issues,'订单合并','','合并后字段缺失',str(absent_after_merge)+'；请查看原始订单表头')
        return None
    for col in ('姓名','电话','地址','数量','商品规格'):
        merged[col]=merged[col].map(clean)
    # 联系信息允许历史样本缺失：先做订单匹配和商品统计，再单独阻断模板输出。
    merged['数量']=pd.to_numeric(merged['数量'],errors='coerce')
    if merged['数量'].isna().any() or (merged['数量']<=0).any() or (merged['数量']%1!=0).any():
        for oid in merged.loc[merged['数量'].isna()|(merged['数量']<=0)|(merged['数量']%1!=0),'订单号'].unique():
            review_record(issues,'订单合并',oid,'数量错误','须为正整数')
    if '订单状态' in right and '支付状态' in right:
        bad=right[~right['订单状态'].map(clean).eq('待发货')|~right['支付状态'].map(clean).eq('支付成功')]
        for oid in bad['订单号'].unique():review_record(issues,'订单合并',oid,'订单状态/支付状态不符合筛选','需待发货、支付成功')
    if issues:return None
    merged['数量']=merged['数量'].astype(int)
    # 同一个订单一份运单的假设；多商品订单不自动合并，不等于不能发，只是需要人工判定打包方式。
    counts=merged['订单号'].value_counts()
    for oid in counts[counts>1].index:
        review_record(issues,'订单合并',oid,'同一订单多商品','门店/云仓可能多包裹，需确认实际分包与匹配键')
    for oid,g in merged.groupby('订单号'):
        if g[['姓名','电话','地址']].drop_duplicates().shape[0]>1:
            review_record(issues,'订单合并',oid,'收件信息冲突','同一订单不同收件信息')
    merged['渠道']=merged['商品名称'].map(lambda x:'云仓' if x in CLOUD_NAMES else '门店')
    merged['规格数量']=merged['商品规格'].map(unit_count)
    merged['实际数量']=merged['数量']*merged['规格数量']
    merged['是否偏远地区']=merged['地址'].map(lambda x:any(term in x for term in SENSITIVE_AREA))
    merged['是否积分兑换']=merged['商品名称'].str.contains('积分兑换',regex=False,na=False)
    merged['是否补运费待确认']=merged['是否偏远地区']&merged['是否积分兑换']
    merged['是否港澳台']=merged['地址'].map(lambda x:any(term in x for term in HMT_AREA))
    for _,r in merged[merged['是否港澳台']].iterrows():
        review_record(issues,'订单检查',r['订单号'],'港澳台地址','当前流程不支持常规发货')
    for oid,g in merged.groupby('订单号'):
        if g['渠道'].nunique()>1:review_record(issues,'订单检查',oid,'同单混合门店/云仓','需确认拆包与上传逻辑')
    if not merged.empty:
        missing_names=set(merged['商品名称'])-set(SUMMARY_NAMES)-set(CLOUD_NAMES)
        # 新品仍按旧规则默认门店；仅列入商品复核清单提示，不拦截全部生产数据。
        if missing_names:
            dest=p['01']/'合并结果'/'新品及渠道人工复核.xlsx'
            save(merged[merged['商品名称'].isin(missing_names)][['订单号','商品名称','渠道','商品规格','数量']],dest)
            msg('  注意：存在未收录在商品统计名称映射中的商品，请核对是否应发云仓。')
    out=p['01']/'合并结果'
    save(left.drop(columns='__商品键',errors='ignore'),out/'订单列表_合并.xlsx')
    save(right.drop(columns='__商品键',errors='ignore'),out/'订单_合并.xlsx')
    save(merged.drop(columns='__商品键',errors='ignore'),out/'一键生成py.xlsx')
    summary=merged.copy()
    summary['统计商品名称']=summary['商品名称'].replace(SUMMARY_NAMES)
    group=summary.groupby('统计商品名称',as_index=False)['实际数量'].sum()
    group=group[['实际数量','统计商品名称']].rename(columns={'统计商品名称':'商品名称'})
    group['__sort']=group['商品名称'].str.extract(r'^(\d+)')[0].fillna('999').astype(int)
    group=group.sort_values(['__sort','商品名称']).drop(columns='__sort')
    save(group,Path(raw_folder)/'商品统计表.xlsx')
    merged, supplement = supplement_recipient(merged, out, issues)
    save(merged.drop(columns='__商品键',errors='ignore'),out/'一键生成py.xlsx')
    # 当源文件中缺手机号/地址时，商品统计仍可生成，但绝不写中通待导入模板。
    for col in ('姓名','电话','地址'):
        missing_ids=merged.loc[merged[col].eq(''),'订单号'].unique()
        for oid in missing_ids:
            review_record(issues,'收件信息',oid,col+'缺失','请在合并结果/收件信息补录.xlsx里填入原系统核实的信息')
    # 即使模板尚不能生成，也能看到统计及分流汇总；不能将预览数据直接用于发货。
    channels=merged.groupby('渠道',as_index=False).agg(订单条数=('订单号','size'),订单数=('订单号','nunique'),实际件数=('实际数量','sum'))
    save(channels,out/'分流统计_预览.xlsx')
    if issues:
        invalidate_imports(p,dt)
        msg('  已完成商品统计和渠道预览；存在收件信息或订单异常，暂不生成中通正式模板。')
        return merged
    self_df=merged[merged['渠道'].eq('门店')].copy()
    cloud_df=merged[merged['渠道'].eq('云仓')].copy()
    self_df['品名']=self_df.apply(lambda r:f"【{r['实际数量']}】 * {self_goods(r['商品名称'],r['商品规格'])}",axis=1)
    cloud_df['仓库货号']=cloud_df['商品名称'].map(cloud_goods)
    for _,r in cloud_df[cloud_df['仓库货号'].eq('')].iterrows():
        review_record(issues,'云仓货号',r['订单号'],'云仓商品缺货号','禁止发送云仓，请完善映射')
    cloud_df['品名']=cloud_df.apply(lambda r:f"【{r['实际数量']}】 * {r['仓库货号']}",axis=1)
    # V3.3：只生成云仓文件和门店审核所需的基础数据；门店必须等用户人工审核确认。
    # 「补运费清单_需人工确认」「补运费人工确认记录」这两份文件不再生成。
    # 云仓的08目录由 prepare_cloud_from_audit 输出：一份原始审核表，一份中通模板。
    # 不在此阶段写云仓模板，否则人工审核修改F/G/H不生效。
    if not issues:
        sig_file.write_text(json.dumps({'signature':source_sig,'source_count':len(files),'date':str(dt)},ensure_ascii=False,indent=2),encoding='utf-8')
    msg(f'  阶段1完成：门店待人工审核{len(self_df)}行、云仓{len(cloud_df)}行。')
    return merged

def find_back_file(folder,flags):
    files=sorted([p for p in folder.glob('*.xls*') if p.suffix.lower() in ('.xlsx','.xls') and not p.name.startswith('~$') and not any(s in p.name for s in flags)])
    if len(files)==0:return None
    if len(files)>1:
        msg(f'  提示：{folder} 有多个回传文件，将合并读取（请确认不是重复导出）：')
        for f in files:msg('    '+f.name)
    return files

# 物流系统常见商品字段；保留商品信息，不能像V4.2那样在读取时丢弃。
WAYBILL_GOODS_COLUMNS = ('品名','商品名称','物品名称','货品名称','产品名称',
                         '订单商品','商品信息','物品信息','物品详情','商品详情',
                         '商品规格名称','SKU名称','货品','商品')


def load_waybills(files):
    frames=[]
    for file in files:
        df=pd.read_excel(file,dtype=str,keep_default_na=False)
        df=aliases(canonical_cols(df),WAYBILL_ALIASES)
        required=['物流单号','姓名','电话']
        missing=[c for c in required if c not in df]
        if missing:raise ValueError(f'{file.name} 缺列：{missing}，实际列名：{list(df.columns)}')
        if '订单编号' not in df:df['订单编号']=''
        product_columns=[c for c in WAYBILL_GOODS_COLUMNS if c in df]
        if product_columns:
            df['回传品名']=df[product_columns].apply(
                lambda r:next((clean(v) for v in r if clean(v)),''),axis=1)
        else:
            df['回传品名']=''
            msg(f'  [提示] {file.name} 没有商品/品名字段；将按手机号稳定分组并保留组内行序；逐单对应依据是两边组内顺序相同。')
        df=df[required+['订单编号','回传品名']].copy()
        # 默认首行为表头：保留物流回传原始 Excel 行号，方便逐项核查。
        df['__Excel行号']=range(2,len(df)+2)
        df['__源文件']=file.name
        frames.append(df)
    bill=pd.concat(frames,ignore_index=True)
    for c in ('物流单号','姓名','电话','订单编号','回传品名'):
        bill[c]=bill[c].map(clean)
    return bill


def recipient_key(s):
    s=unicodedata.normalize('NFKC',clean(s))
    s=re.sub(r'^(?:收件人|收货人)\s*[:：]\s*','',s)
    # Excel/WPS 粘贴时可能包含肉眼看不见的零宽字符。
    s=re.sub(r'[\u200b-\u200d\ufeff\u2060]','',s)
    return re.sub(r'\s+','',s).casefold()


def phone_key(s):
    x=unicodedata.normalize('NFKC',clean(s))
    x=re.sub(r'[\s\-()\u200b-\u200d\ufeff\u2060]','',x)
    if x.startswith('+86') and len(x)==14:x=x[3:]
    elif x.startswith('86') and len(x)==13:x=x[2:]
    return x


def product_quantity_and_token(s):
    """只读取明确的【数量】*品名格式，不把5件套中的5当作商品件数。"""
    x=unicodedata.normalize('NFKC',clean(s)).casefold()
    x=re.sub(r'\s+','',x)
    qty=None
    m=re.match(r'^\[(\d+)\][*x×:：-]?',x)
    if m:
        qty=int(m.group(1))
        x=x[m.end():]
    x=x.replace('＊','*')
    # 只统一符号/空白，不采用不安全的模糊包含匹配。
    return qty,x


def expected_products(row, channel):
    name=clean(row['商品名称']);spec=clean(row.get('商品规格',''))
    if channel=='门店':
        shipping=self_goods(name,spec)
    else:
        shipping=cloud_goods(name)
    q=clean(row.get('实际数量',''))
    candidates={name,shipping, f'【{q}】 * {shipping}' if q and shipping else ''}
    tokens={product_quantity_and_token(x)[1] for x in candidates if x}
    tokens.discard('')
    return tokens


def goods_agree(bill_goods, order_row, channel):
    if not clean(bill_goods):return None
    qty,token=product_quantity_and_token(bill_goods)
    if token not in expected_products(order_row,channel):return False
    want=clean(order_row.get('实际数量',''))
    return qty is None or (want.isdigit() and int(want)==qty)


def load_product_confirmations(path, issues, channel):
    """可选人工补充品名或明确指定订单号；不覆盖物流原件。"""
    if not path.exists():return {}
    try:
        tab=pd.read_excel(path,dtype=str,keep_default_na=False)
        required=['物流单号','补充商品名称','指定订单号']
        if any(c not in tab.columns for c in required):
            raise ValueError('商品核对表缺少'+str([c for c in required if c not in tab]))
        tab['物流单号']=tab['物流单号'].map(clean)
        if tab['物流单号'].eq('').any() or tab['物流单号'].duplicated().any():
            raise ValueError('商品核对表的物流单号必须非空且唯一')
        return {clean(r['物流单号']): {'品名':clean(r['补充商品名称']),
                                   '订单号':clean(r['指定订单号'])}
                for _,r in tab.iterrows()}
    except Exception as exc:
        review_record(issues,channel+'运单','','无法读取运单商品核对表',exc)
        msg('  [暂停] 运单商品补充表有错误：'+str(exc))
        return None


def prepare_product_confirmation(path, bills, unresolved, channel, expected_orders):
    """只在01合并结果生成，不在05/09多生成Excel。永不覆盖用户填过的表。"""
    if not unresolved or path.exists():return
    seen=set();arr=[]
    for r in unresolved:
        no=clean(r['物流单号'])
        if not no or no in seen:continue
        seen.add(no)
        possible=[o for o in expected_orders if
                  recipient_key(o['姓名'])==recipient_key(r['姓名']) and
                  phone_key(o['电话'])==phone_key(r['电话'])]
        suggestions='；'.join(f"{clean(o['订单号'])} / {clean(o['商品名称'])}" for o in possible)
        arr.append({'物流单号':excel_safe(no),'回传收件人':clean(r['姓名']),
                    '回传电话':clean(r['电话']),'物流回传商品':clean(r['回传品名']),
                    '候选订单号及商品（仅参考）':suggestions,
                    '补充商品名称':'','指定订单号':'',
                    '使用说明':'填写核实的商品名；同人同商品多单务必填准确订单号；不可凭排序猜测'})
    if arr:
        save(pd.DataFrame(arr),path)
        msg(f'  [需补商品] 已生成 {path.name}，填写后保存、按回车重新匹配。')


def archive_invalid_import(file, p):
    """不可再直接使用的旧导入文件备份到01目录，避免直接删除用户资料。"""
    if not file.exists():
        return
    target_folder=p['01']/'合并结果'/'旧版不完整导入备份'
    target_folder.mkdir(parents=True,exist_ok=True)
    target=target_folder/file.name
    if target.exists():
        target=target_folder/(file.stem+'_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f')+file.suffix)
    shutil.move(str(file),str(target))
    msg('  已将旧版不完整导入文件移入备份：'+str(target))


def remove_old_manual_files(p):
    """05/09不能出现程序生成的人工匹配附加表；历史文件仅备份移走。"""
    backup=p['01']/'合并结果'/'旧版人工匹配文件备份'
    for channel,code in (('门店','05'),('云仓','09')):
        legacy=p[code]/('人工匹配_'+channel+'.xlsx')
        if legacy.is_file():
            backup.mkdir(parents=True,exist_ok=True)
            dest=backup/legacy.name
            if dest.exists():
                dest=backup/(legacy.stem+'_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f')+legacy.suffix)
            shutil.move(str(legacy),str(dest))
            msg(f'  已将旧版多余人工匹配文件移至备份：{dest}')


def audit_fee_orders(p,dt,expected):
    """按照审核表H列确认的当前分组，而非原始地址推断的初始值。"""
    dest=audit_file(p,dt,'门店')
    sheet=pd.read_excel(dest,sheet_name='门店原始订单',dtype=str,keep_default_na=False)
    sheet=canonical_cols(sheet)
    if '订单号' not in sheet or '是否需要补运费' not in sheet:
        raise ValueError('门店审核表缺少订单号或是否需要补运费，无法安全分组')
    sheet['订单号']=sheet['订单号'].map(clean)
    if sheet['订单号'].duplicated().any() or set(sheet['订单号'])!=set(expected['订单号']):
        raise ValueError('门店审核表订单号数量与原始订单不一致，请先核查')
    return {clean(r['订单号']):parse_fee_cell(r['是否需要补运费'],r['订单号'])[0]
            for _,r in sheet.iterrows()}


def save_reconcile_report(p,dt,channel,summary,expected_detail,unlinked,warnings=None):
    """只在01/合并结果记录匹配证据与行序/姓名提醒，05/09不创建附加文件。"""
    dest=p['01']/'合并结果'/f'{dt.month}.{dt.day}_运单对账_{channel}.xlsx'
    dest.parent.mkdir(parents=True,exist_ok=True)
    warning_cols=['订单号','物流单号','电话尾号','原单收件人','回传收件人','组内序号','中通模板Excel行号','提醒','回传来源','回传Excel行号']
    with pd.ExcelWriter(dest,engine='openpyxl') as writer:
        pd.DataFrame(summary,columns=['核对项目','数量']).to_excel(writer,index=False,sheet_name='数量核对')
        pd.DataFrame(expected_detail).to_excel(writer,index=False,sheet_name='原始订单逐条核对')
        anomaly_cols=['回传来源','回传Excel行号','物流单号','回传订单号','回传姓名','回传电话','电话后四位','回传品名',
                      '原单候选订单号','原单候选姓名','原单候选行号','原单候选数','异常原因',
                      '回传姓名标准化','原单姓名标准化','定位提示']
        pd.DataFrame(unlinked,columns=anomaly_cols).to_excel(writer,index=False,sheet_name='回传异常明细')
        pd.DataFrame(warnings or [],columns=warning_cols).to_excel(writer,index=False,sheet_name='行序匹配提醒')
    msg('  [对账明细] '+str(dest))


def save_row_order_review(path, proposed, issues, label):
    """产出/读取同收件人的组内行序暂配对账；只有逐条标记「确认」才可正式导入。

    历史复核表有用户填写内容，不能覆盖。如配对关系发生变化，先备份旧表再新建。
    """
    cols=['订单号','物流单号','收件人姓名','收件人电话','商品名称','配对方式','配对说明','核对结果']
    new_rows=[]
    for r in proposed:
        new_rows.append({c:excel_safe(r.get(c,'')) for c in cols})
    if not new_rows:
        return set()
    dest=Path(path)
    expected={(clean(r['订单号']),clean(r['物流单号'])) for r in new_rows}
    if dest.exists():
        try:
            old=pd.read_excel(dest,dtype=str,keep_default_na=False)
            if not all(c in old.columns for c in cols):
                raise ValueError('缺少必备核对列')
            pairs=[(clean(r['订单号']),clean(r['物流单号'])) for _,r in old.iterrows()]
            if len(pairs)!=len(set(pairs)) or set(pairs)!=expected:
                raise ValueError('行序候选关系或数量已经变化')
            bad=old.loc[~old['核对结果'].map(clean).isin(['','确认','待复核'])]
            if not bad.empty:
                raise ValueError('「核对结果」只能填「确认」或留空')
            confirmed={(clean(r['订单号']),clean(r['物流单号']))
                       for _,r in old.iterrows() if clean(r['核对结果'])=='确认'}
            return confirmed
        except Exception as e:
            folder=dest.parent/'旧版行序复核表备份'
            folder.mkdir(parents=True,exist_ok=True)
            back=folder/(dest.stem+'_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f')+'.xlsx')
            shutil.copy2(dest,back)
            msg(f'  [行序关系变化] 旧复核表已备份：{back.name}；原因：{e}')
            review_record(issues,label+'运单','','旧行序复核关系变化，确认状态失效',str(e))
    save(pd.DataFrame(new_rows,columns=cols),dest)
    msg('  [待逐单复核] 行序候选表：'+str(dest))
    return set()


def row_order_pair_candidates(pending, receiver_map, exps, matched, fee_receivers, label):
    """只在同姓名+同电话组内按原文件出现顺序暂配，绝不跨组/按全局行号。"""
    grouped=defaultdict(list)
    for r in pending:
        grouped[(recipient_key(r['姓名']),phone_key(r['电话']))].append(r)
    proposals=[];paired_bill=set();notes=[]
    for group,rs in grouped.items():
        candidates=[oid for oid in receiver_map.get(group,[]) if oid not in matched]
        if not candidates:
            continue
        if len(candidates)!=len(rs):
            notes.append(f'{group[0]} / 电话尾号{group[1][-4:]}：未匹配订单{len(candidates)}条，回传{len(rs)}行')
            continue
        # 保留订单首次出现在运营明细/生成模板中的顺序；回传按05/09文件行顺序。
        for index,(oid,r) in enumerate(zip(candidates,rs),start=1):
            order=exps[oid]
            description=f'同收件人组内第{index}条；回传不含商品/订单号，行序不可证明真实对应'
            if group in fee_receivers:
                description+='；该收件人还存在人工补运费单'
            proposals.append({'订单号':oid,'物流单号':clean(r['物流单号']),
                '收件人姓名':clean(order['姓名']),'收件人电话':clean(order['电话']),
                '商品名称':clean(order['商品名称']),'配对方式':'同姓名电话组内行序暂配',
                '配对说明':description,'核对结果':''})
            paired_bill.add(clean(r['物流单号']))
    return proposals,paired_bill,notes


def save_review_import(output, warnings, target, channel):
    """新运管两工作表：第一个严格三列；第二个说明暂配/姓名差异。

    不因标记差异而自动生成「核对通过」文件，避免误作已人工确认。
    若人工曾在第二表填确认，不随重跑静默覆盖；旧文件先备份。
    """
    from openpyxl.styles import Font, Alignment
    target.parent.mkdir(parents=True, exist_ok=True)
    risk_cols=['订单号','物流单号','原始订单收件人','中通回传收件人','手机号后四位',
               '物流回传来源','回传Excel行号','中通模板Excel行号','匹配方式/风险','需要核对的内容','人工核对结果']
    rows=[]
    for w in warnings:
        rows.append({
            '订单号': excel_safe(w['订单号']), '物流单号': excel_safe(w['物流单号']),
            '原始订单收件人': excel_safe(w['原单收件人']),
            '中通回传收件人': excel_safe(w['回传收件人']),
            '手机号后四位': w['电话尾号'],
            '物流回传来源': w.get('回传来源',''),
            '回传Excel行号': w.get('回传Excel行号',''),
            '中通模板Excel行号': w.get('中通模板Excel行号',''),
            '匹配方式/风险': ('姓名不一致，按唯一手机号暂配' if '手机号两侧均唯一' in w['提醒'] else '同手机号多订单按行序暂配'),
            '需要核对的内容': w['提醒'],
            '人工核对结果': ''
        })
    risk_df=pd.DataFrame(rows,columns=risk_cols)
    if target.exists():
        try:
            existing=pd.read_excel(target,sheet_name=None,dtype=str,keep_default_na=False)
            prior=existing.get('新运管导入')
            prior_notes=existing.get('人工复核说明')
            df_out=pd.DataFrame(output,columns=OUTPUT_COLS)
            if prior is not None and prior_notes is not None:
                same_import=(prior[OUTPUT_COLS].astype(str).reset_index(drop=True).equals(
                    df_out[OUTPUT_COLS].astype(str).reset_index(drop=True)))
                compare_cols=[x for x in risk_cols if x!='人工核对结果']
                same_notes=all(x in prior_notes for x in compare_cols) and (prior_notes[compare_cols].astype(str).reset_index(drop=True).equals(
                    risk_df[compare_cols].astype(str).reset_index(drop=True)))
                if same_import and same_notes:
                    msg('  [保留已有文件] 待人工复核文件配对关系未变化，不覆盖人工填写内容：'+str(target))
                    return
        except Exception:
            pass
        # 文件可能已被人工修改；先备份，避免重跑丢失复核记录。
        folder=target.parent/'旧版复核文件备份'
        folder.mkdir(parents=True,exist_ok=True)
        dest=folder/(target.stem+'_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f')+target.suffix)
        shutil.copy2(target,dest)
        msg('  [配对变化] 旧版待复核文件已备份：'+str(dest))
    temp=target.with_name(target.stem+'_临时生成'+target.suffix)
    try:
        with pd.ExcelWriter(temp,engine='openpyxl') as writer:
            pd.DataFrame(output,columns=OUTPUT_COLS).to_excel(writer,index=False,sheet_name='新运管导入')
            risk_df.to_excel(writer,index=False,sheet_name='人工复核说明')
            sheet=writer.sheets['新运管导入']
            sheet.freeze_panes='A2'
            sheet.column_dimensions['A'].width=18
            sheet.column_dimensions['B'].width=26
            sheet.column_dimensions['C'].width=28
            note=writer.sheets['人工复核说明']
            note.freeze_panes='A2'
            note.auto_filter.ref=note.dimensions
            for col,width in {'A':28,'B':26,'C':24,'D':24,'E':18,'F':30,'G':18,'H':22,'I':34,'J':76,'K':18}.items():
                note.column_dimensions[col].width=width
            for cell in note[1]:
                cell.font=Font(bold=True)
                cell.alignment=Alignment(wrap_text=True)
        temp.replace(target)
    finally:
        if temp.exists():temp.unlink()
    msg(f'  [已生成·待人工复核] {len(output)}条{channel}三列数据，附{len(warnings)}项风险说明：{target}')
    msg('  核对第二Sheet并在「人工核对结果」填写「确认」；保存后按回车生成只含三列的单Sheet正式导入文件。')


def review_confirmation_count(path, expected_rows):
    """仅允许相同订单号/物流单号关系沿用人工确认；不凭空认定风险已排除。"""
    try:
        notes=pd.read_excel(path,sheet_name='人工复核说明',dtype=str,keep_default_na=False)
        if len(notes)!=expected_rows:return 0
        states=notes['人工核对结果'].map(clean) if '人工核对结果' in notes else pd.Series([],dtype=str)
        return int(states.isin(['确认','已确认']).sum())
    except Exception:
        return 0


def reconcile(p,dt,channel,rows,files,issues):
    """V4.10 继承原规则，并把手机号唯一而姓名不同的记录列入人工复核。

    手机号只出现一次：先按手机号唯一性暂配，姓名差异必须人工核对。
    手机号出现多次：先恢复制单模板实际行序，再对两侧手机号稳定分组，按组内行序暂配。
    重复手机号组内即使收件人姓名有差异也只作提醒，不因姓名差异整组丢单。
    仍假设中通回传保留每个手机号组内的模板顺序；无法独立证明时继续人工复核。
    """
    label=channel
    msg(f'[运单核对] {label}回传匹配（V4.11：以实际中通模板行序为准 + 人工复核）')
    original=rows[rows['渠道'].eq(channel)].copy().reset_index(drop=True)
    if original.empty:
        msg('  该渠道没有原始订单，跳过。')
        return
    if original['订单号'].map(clean).duplicated().any():
        review_record(issues,label+'运单','','订单号在本渠道重复','请先确认一单多商品或重复导出，不按行序强行分配')
        msg('  [暂停] 原始订单号不唯一。')
        return
    try:
        bills=load_waybills(files).reset_index(drop=True)
        fees=audit_fee_orders(p,dt,original) if channel=='门店' else {}
    except Exception as exc:
        review_record(issues,label+'运单','','回传或审核表读取失败',exc)
        msg('  [暂停] '+str(exc))
        return
    fee_ids={clean(oid) for oid,yes in fees.items() if yes}
    if channel=='门店':
        expected=original[~original['订单号'].map(clean).isin(fee_ids)].copy().reset_index(drop=True)
        msg(f'  门店原始{len(original)}条，其中补运费{len(fee_ids)}条手工发货（不导入07），普通单{len(expected)}条。')
        archive_invalid_import(p['07']/f'{dt.month}.{dt.day}_新运管导入_门店补运费_核对通过.xlsx',p)
    else:
        expected=original.copy().reset_index(drop=True)
    msg(f'  本批物流回传{len(bills)}行；本渠道应导入{len(expected)}条。')
    try:
        expected=expected_in_template_order(expected,p,dt,channel)
        msg(f'  [{label}对账顺序] 已采用中通实际发货模板行序，而非订单合并表行序。')
    except Exception as e:
        review_record(issues,label+'运单','','模板顺序映射丢失或无效',str(e))
        msg(f'  [暂停] {e}')
        return

    # 实际发货模板行序已从映射恢复；稳定排序只改变手机号组排列，不改变组内制单顺序。
    ex=expected.copy()
    ex['__电话键']=ex['电话'].map(phone_key)
    ex['__姓名键']=ex['姓名'].map(recipient_key)
    ex['__输入行号']=ex['__模板Excel行号']
    bl=bills.copy()
    bl['__电话键']=bl['电话'].map(phone_key)
    bl['__姓名键']=bl['姓名'].map(recipient_key)
    bl['__输入行号']=range(1,len(bl)+1)
    # 使用字符串而非浮点数字作为手机号/运单号；stable保持每个号码下的原有顺序。
    ex=ex.sort_values('__电话键',kind='stable').reset_index(drop=True)
    bl=bl.sort_values('__电话键',kind='stable').reset_index(drop=True)
    ex_groups={p:g for p,g in ex.groupby('__电话键',sort=False)}
    bill_groups={p:g for p,g in bl.groupby('__电话键',sort=False)}
    bad_bill_nos=set(bl.loc[bl['物流单号'].map(clean).duplicated(keep=False),'物流单号'].map(clean))-{''}

    matched={}; methods={}; sequence_index={}; problems=[]; warnings=[]; touched=set()
    failed_by_order=defaultdict(list)
    def problem(r,reason):
        phone=phone_key(r.get('电话',''))
        candidate_group=ex_groups.get(phone)
        candidates=candidate_group.to_dict('records') if candidate_group is not None else []
        rowno=r.get('__Excel行号','?')
        src=clean(r.get('__源文件',''))
        names='、'.join(clean(x.get('姓名','')) for x in candidates)
        order_nos='、'.join(clean(x.get('订单号','')) for x in candidates)
        source_rows='、'.join(str(x.get('__输入行号','')) for x in candidates)
        standard_names='、'.join(recipient_key(x.get('姓名','')) for x in candidates)
        diag='手机号找不到本批原始订单' if not candidates else ('有原单候选，请核对标准化姓名及回传订单号/物流单号')
        problems.append({'回传来源':src,'回传Excel行号':rowno,
                         '物流单号':clean(r.get('物流单号','')),
                         '回传订单号':clean(r.get('订单编号','')),'回传姓名':clean(r.get('姓名','')),
                         '回传电话':clean(r.get('电话','')),'电话后四位':phone[-4:],
                         '回传品名':clean(r.get('回传品名','')),'原单候选订单号':order_nos,
                         '原单候选姓名':names,'原单候选行号':source_rows,
                         '原单候选数':len(candidates),'异常原因':reason,
                         '回传姓名标准化':recipient_key(r.get('姓名','')),
                         '原单姓名标准化':standard_names,'定位提示':diag})
        for x in candidates:
            failed_by_order[clean(x.get('订单号',''))].append(reason)
        # 控制台只显示手机号尾号，完整姓名、电话和订单号保存在本地对账表中。
        msg(f'  [异常定位] {src} 第{rowno}行，手机号尾号{phone[-4:] or "未知"}：{reason}；原单候选{len(candidates)}条')
        review_record(issues,label+'运单',clean(r.get('订单编号','')),reason,f'{src} 第{rowno}行')
    def warn(o,r,pos,reason):
        warnings.append({'订单号':clean(o['订单号']),'物流单号':clean(r['物流单号']),
                         '电话尾号':phone_key(r['电话'])[-4:],
                         '原单收件人':clean(o['姓名']),'回传收件人':clean(r['姓名']),
                         '组内序号':pos,'中通模板Excel行号':o.get('__模板Excel行号',''),'提醒':reason,
                         '回传来源':clean(r.get('__源文件','')),
                         '回传Excel行号':r.get('__Excel行号','')})

    # 错误组不尝试部分填充：即使其他组能匹配，仍不生成不完整的07文件。
    for ph in dict.fromkeys(list(ex_groups)+list(bill_groups)):
        os=list(ex_groups.get(ph,pd.DataFrame()).to_dict('records')) if ph in ex_groups else []
        rs=list(bill_groups.get(ph,pd.DataFrame()).to_dict('records')) if ph in bill_groups else []
        if not ph or unsafe_id(ph):
            for r in rs:problem(r,'手机号为空或存在科学计数法，无法配对')
            if os:review_record(issues,label+'运单','','原始订单存在空/不可信手机号',f'数量{len(os)}')
            continue
        if len(os)!=len(rs):
            reason=f'该手机号组数量不一致：待导入订单{len(os)}条、物流回传{len(rs)}条，禁止行序配对'
            for r in rs:problem(r,reason)
            if os:
                for o in os:failed_by_order[clean(o.get('订单号',''))].append(reason)
                review_record(issues,label+'运单','','手机号组数量不一致',f'电话尾号{ph[-4:]}：{reason}')
                msg(f'  [异常定位] 原始订单手机号尾号{ph[-4:]}：{reason}；原始行号' +
                    '、'.join(str(o.get('__输入行号','')) for o in os))
            continue
        for pos,(o,r) in enumerate(zip(os,rs),start=1):
            oid=clean(o['订单号']);no=clean(r['物流单号'])
            if not no or unsafe_id(no):
                problem(r,'物流单号缺失/科学计数法可能导致精度丢失');continue
            if no in bad_bill_nos:
                problem(r,'物流单号在回传中重复，请先检查是否合包或重复导出');continue
            return_oid=clean(r.get('订单编号',''))
            if return_oid and return_oid!=oid:
                problem(r,f'回传自带订单编号与按行序对应的原订单号不一致：{return_oid}；禁止忽略明确订单号');continue
            if return_oid and unsafe_id(return_oid):
                problem(r,'回传订单号是科学计数法，可能丢失精度');continue
            # V4.10：唯一手机号存在姓名差异时，只在两侧各唯一且物流/订单号没有硬冲突的前提下暂配。
            # 不把名字中的问号或间隔号直接视为相等：原样保留两侧姓名，写入第二工作表。
            name_diff = (o['__姓名键'] != r['__姓名键'])
            how=('手机号唯一但姓名不一致：暂配待人工核对' if name_diff else '姓名+手机号唯一匹配') if len(os)==1 else '重复手机号组内稳定行序匹配'
            matched[oid]=no
            methods[oid]=how
            sequence_index[oid]=pos
            touched.add(no)
            if len(os)==1 and name_diff:
                warn(o,r,pos,'手机号两侧均唯一，但姓名不一致；按手机号暂配。请对照物流系统核实姓名符号与真实收件人')
                msg(f'  [姓名差异待核对] {clean(r.get("__源文件",""))} 第{r.get("__Excel行号","?")}行，手机号尾号{ph[-4:]}；已暂配，详见07文件第二工作表')
            elif len(os)>1 and name_diff:
                warn(o,r,pos,'手机号重复、按原行序配对，但收件人姓名不同：请重点复核；不自动改变组内顺序')
            elif len(os)>1:
                warn(o,r,pos,'手机号重复，按两边组内原行序配对：须确保中通未改变组内顺序')

    # 07仅生成一份完整的三列文件，不输出含缺项的假完成文件。
    detail=[]
    for _,r in original.iterrows():
        oid=clean(r['订单号']);is_fee=label=='门店' and oid in fee_ids
        ship_name=self_goods(r['商品名称'],clean(r.get('商品规格',''))) if label=='门店' else cloud_goods(r['商品名称'])
        qty=clean(r.get('实际数量',''))
        if is_fee:
            status='补运费手工发货，不导入07';why='';bill='';how='手动发货'
        elif oid in matched:
            status='按所选规则已配对';why='同手机号多订单依赖两边组内顺序一致' if sequence_index[oid] and methods[oid].startswith('重复') else ''
            bill=matched[oid];how=methods[oid]
        else:
            status='未匹配';why='；'.join(dict.fromkeys(failed_by_order.get(oid,[]))) or '对应的手机号组或物流单号异常，详见回传异常明细';bill='';how=''
        detail.append({'渠道':label,'订单分组':'补运费手工处理' if is_fee else '正常订单',
                       '订单号':oid,'商品名称':clean(r['商品名称']),
                       '预期发货品名':f'【{qty}】 * {ship_name}',
                       '收件人姓名':clean(r['姓名']),'收件人电话':clean(r['电话']),
                       '匹配状态':status,'物流单号':bill,'未匹配原因':why,'匹配方式':how,
                       '组内序号':sequence_index.get(oid,''),
                       '中通模板Excel行号':next((x['__模板Excel行号'] for _,x in expected.iterrows() if clean(x['订单号'])==oid),'')})
    summary=[('原始订单总数',len(original)),('补运费手工发货（不导入）',len(fee_ids)),
             ('应导入订单数',len(expected)),('回传文件个数',len(files)),('回传总行数',len(bl)),
             ('已按规则配对订单数',len(matched)),('尚未配对订单数',len(expected)-len(matched)),
             ('手机号重复的行序配对条数',sum(1 for x in methods.values() if x.startswith('重复'))),
             ('行序/姓名风险提醒数',len(warnings)),('回传异常数',len(problems))]
    save_reconcile_report(p,dt,label,summary,detail,problems,warnings)
    msg(f'  {label}：按规则配对{len(matched)}/{len(expected)}；重复手机号行序配对{sum(1 for x in methods.values() if x.startswith("重复"))}条，提醒{len(warnings)}条。')
    if problems:
        msg(f'  [异常] {len(problems)}条回传需要核查；上面已显示具体回传行号，完整原因见对账表「回传异常明细」。')

    dest_ok=p['07']/f'{dt.month}.{dt.day}_新运管导入_{label}_核对通过.xlsx'
    dest_review=p['07']/f'{dt.month}.{dt.day}_新运管导入_{label}_待人工复核.xlsx'
    blockers=[x for x in issues if x['阶段'] in ('订单检查','订单合并')]
    ready=(len(expected)>0 and len(expected)==len(bills) and len(matched)==len(expected)
           and len(set(matched.values()))==len(matched) and not problems and not blockers)
    if ready:
        output=[{'物流公司名称':'中通快递','物流单号':excel_safe(matched[oid]),
                 '订单编号':excel_safe(oid)} for oid in expected['订单号'].map(clean)]
        if warnings:
            # 第一轮始终生成带风险Sheet的候选文件；核对后可在第二Sheet的末列填「确认」。
            # 再按回车生成仅一张工作表的正式导入文件，避免新运管拒绝多Sheet。
            save_review_import(output,warnings,dest_review,label)
            confirmed=review_confirmation_count(dest_review,len(warnings))
            if confirmed==len(warnings):
                save(pd.DataFrame(output,columns=OUTPUT_COLS),dest_ok)
                msg(f'  [人工已确认] {len(warnings)}/{len(warnings)}项；另生成纯三列正式导入文件：{dest_ok}')
            else:
                archive_invalid_import(dest_ok,p)
                msg(f'  [待人工确认] 已确认{confirmed}/{len(warnings)}项。请在第二Sheet「人工核对结果」列核实后填「确认」，保存并按回车生成可直接导入的单Sheet文件。')
        else:
            archive_invalid_import(dest_review,p)
            save(pd.DataFrame(output,columns=OUTPUT_COLS),dest_ok)
            msg(f'  [已生成] {len(output)}条{label}三列导入文件，未发现姓名差异或行序暂配风险。')
    else:
        archive_invalid_import(dest_ok,p)
        archive_invalid_import(dest_review,p)
        msg('  [未完成] 数量或关键字段核对未通过，已撤下可能误配的07导入文件。')
    archive_invalid_import(p['07']/f'{dt.month}.{dt.day}_{label}_匹配成功明细_禁止直接导入.xlsx',p)
    # 升级V4.4留下的“逐行手动确认”不再是必要步骤；为避免误读，移动到01备份目录。
    old_review=p['01']/'合并结果'/f'{dt.month}.{dt.day}_行序配对人工复核_{label}.xlsx'
    if old_review.exists():
        backup=p['01']/'合并结果'/'旧版行序复核表备份'
        backup.mkdir(parents=True,exist_ok=True)
        dest=backup/old_review.name
        if dest.exists():dest=backup/(old_review.stem+'_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f')+'.xlsx')
        shutil.move(str(old_review),str(dest))
        msg('  旧版需逐笔确认的复核表已备份，不再要求填写：'+old_review.name)

def get_raw_folder(p,root,dt,user_arg):
    if user_arg:
        path=Path(user_arg.strip().strip('"').strip("'"))
        if not path.is_dir():raise FileNotFoundError(f'输入文件夹不存在：{path}')
        return path
    return p['01']

def run(root,dt,order_folder=None):
    global DETAIL_LOG_FILE
    root=Path(root).expanduser().resolve()
    root.mkdir(parents=True,exist_ok=True)
    p=paths(root,dt)
    DETAIL_LOG_FILE = p['01']/'合并结果'/'运行日志.txt'
    DETAIL_LOG_FILE.parent.mkdir(parents=True,exist_ok=True)
    msg('\n========== 一体化发货助手 / '+str(dt)+' ==========')
    move_old_bonus_excel(p,dt)
    remove_old_manual_files(p)
    for code in CATEGORIES:msg(f' {code}：{p[code]}')
    raw_folder=get_raw_folder(p,root,dt,order_folder)
    msg('订单来源：'+str(raw_folder))
    issues=[]
    merged=stage_prepare(raw_folder,p,dt,root,issues)
    if merged is None:
        # 如果本次输入不足，不借用旧合并结果进行发货，避免混入历史批次。
        invalidate_imports(p,dt)
        report(issues,p['01']/'合并结果'/'运行异常与提示.xlsx')
        msg('未完成阶段1；阶段2/3不执行。')
        return 1
    blockers={'订单检查','订单合并','订单导出','收件信息','收件补录'}
    if any(i['阶段'] in blockers for i in issues):
        self_ready=cloud_ready=False
        msg('[暂停] 订单基础信息尚未通过验证，只输出统计与异常清单，不生成门店/云仓审核及导入模板。')
    else:
        self_ready=prepare_self_from_audit(merged,p,dt,root,issues)
        cloud_ready=prepare_cloud_from_audit(merged,p,dt,root,issues)
        if self_ready:
            self_ready=ensure_template_order_map(merged,p,dt,'门店',issues)
        if cloud_ready:
            cloud_ready=ensure_template_order_map(merged,p,dt,'云仓',issues)
    for channel,code,exclude in [('门店','05',['人工匹配_']),('云仓','09',['人工匹配_'])]:
        files=find_back_file(p[code],exclude)
        if files:
            for back in files:
                stamp=re.search(r'(20\d{6})',back.stem)
                if stamp and stamp.group(1)!=dt.strftime('%Y%m%d'):
                    msg(f'  [注意文件日期] {channel}回传「{back.name}」包含日期{stamp.group(1)}，与当前批次{dt}不同；请确认不是混入历史回传。')
        if files is None:
            stale=p['07']/f'{dt.month}.{dt.day}_新运管导入_{channel}_核对通过.xlsx'
            if stale.is_file():archive_invalid_import(stale,p)
            if channel=='门店':
                extra=p['07']/f'{dt.month}.{dt.day}_新运管导入_门店补运费_核对通过.xlsx'
                if extra.exists():archive_invalid_import(extra,p)
            msg(f'[等待] {channel}运单文件尚未放入 {p[code]}；放入后按回车继续。')
        elif (channel=='门店' and not self_ready) or (channel=='云仓' and not cloud_ready):
            candidate=p['07']/f'{dt.month}.{dt.day}_新运管导入_{channel}_待人工复核.xlsx'
            if candidate.exists():archive_invalid_import(candidate,p)
            old=p['07']/f'{dt.month}.{dt.day}_新运管导入_{channel}_核对通过.xlsx'
            if old.exists():archive_invalid_import(old,p)
            if channel=='门店':
                extra=p['07']/f'{dt.month}.{dt.day}_新运管导入_门店补运费_核对通过.xlsx'
                if extra.exists():archive_invalid_import(extra,p)
            msg('[等待] '+channel+'审核表/模板未通过检查；程序将继续保护已人工修改的文件。')
            review_path=audit_file(p,dt,channel)
            target_template=(self_template_paths(p,dt)[0] if channel=='门店'
                             else p['08']/f'{dt.month}.{dt.day}.xlsx')
            msg('  核查审核表：'+str(review_path)+(' [存在]' if review_path.is_file() else ' [缺少]'))
            msg('  核查目标模板：'+str(target_template)+(' [存在]' if target_template.is_file() else ' [缺少]'))
            msg('  具体原因请查看：'+str(p['01']/'合并结果'/'运行异常与提示.xlsx'))
        elif any(i['阶段'] in ('订单检查','订单合并','收件信息','收件补录') or (channel=='云仓' and i['阶段']=='云仓货号') for i in issues):
            invalidate_imports(p,dt)
            msg(f'[暂缓] 存在订单/模板基础异常，请先核查，暂不自动匹配{channel}。')
        else:
            # 门店和云仓彼此独立：门店补运费未确认不阻塞云仓。
            reconcile(p,dt,channel,merged,files,issues)
    report(issues,p['01']/'合并结果'/'运行异常与提示.xlsx')
    if issues:msg(f'本次共{len(issues)}项需要核查。请查看运行异常与提示.xlsx')
    else:msg('数据预处理检查无异常；中通实际打印、补运费及新运管最终上传仍需人工确认。')
    msg('  [详情日志] 01日期目录/合并结果/运行日志.txt（需要完整路径、文件名及报错时查看）')
    msg('========== 本次运行结束 ==========')
    return 0 if not issues else 2

def move_old_bonus_excel(p,dt):
    """只迁移旧版本额外的补运费确认表，不删除资料、不触碰3份正式Excel。"""
    from datetime import datetime
    folder=p['03']
    old_names=[folder/f'{dt.month}.{dt.day}_补运费清单_需人工确认.xlsx',
               folder/'补运费人工确认记录.xlsx',
               folder/f'{dt.month}.{dt.day}_补运费人工确认记录.xlsx']
    for old in old_names:
        if old.is_file():
            backup=p['01']/'合并结果'/'旧版多余文件备份'
            backup.mkdir(parents=True,exist_ok=True)
            target=backup/old.name
            if target.exists():
                target=backup/(old.stem+'_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f')+old.suffix)
            shutil.move(str(old),str(target))
            msg('  已将03旧版多余确认表移到01备份（未删除）：'+old.name)


def normalize_old_template_fill(dest,p,channel):
    """兼容已经生成的旧版模板：清除填充色但不改数据和列数，操作前备份。"""
    if not dest.is_file():return
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill
    wb=load_workbook(dest)
    try:
        needs_change=any(cell.fill.patternType not in (None,'none')
                         for ws in wb.worksheets for row in ws for cell in row)
        if not needs_change:return
        backup=p['01']/'合并结果'/'旧模板样式备份'/channel
        backup.mkdir(parents=True,exist_ok=True)
        target=backup/dest.name
        if target.exists():
            target=backup/(dest.stem+'_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f')+dest.suffix)
        shutil.copy2(dest,target)
        clear=PatternFill(fill_type=None)
        for ws in wb.worksheets:
            for row in ws:
                for cell in row:
                    if cell.fill.patternType not in (None,'none'):
                        cell.fill=clear
        tmp=p['01']/'合并结果'/('清除颜色_'+channel+'_'+dest.name)
        wb.save(tmp)
        tmp.replace(dest)
        msg('  已清除已存在模板的填充色（未改内容；原文件备份在01合并结果）：'+dest.name)
    finally:wb.close()


def order_path_settings(folder):
    """从订单日期目录解析发货日期和根目录，例如 .../01xxx/2026/10月/10.30。"""
    raw_path=str(folder).strip().strip('\"').strip("'")
    if not raw_path:
        raise ValueError('请输入8份订单所在的01日期文件夹路径')
    raw=Path(raw_path).expanduser().resolve()
    if not raw.is_dir():
        raise FileNotFoundError('订单文件夹不存在：'+str(raw))
    if raw.name=='订单文件':
        raw=raw.parent
    category=raw.parent.parent.parent
    if not re.match(r'^01(?!\d)', category.name):
        raise ValueError('路径应位于 01分类/年份/月份/日期 中，例如：01一键搞定发货信息/2026/10月/10.30')
    year_name=raw.parent.parent.name
    month_name=raw.parent.name
    day_name=raw.name
    if not re.fullmatch(r'20\d{2}',year_name):
        raise ValueError('日期路径缺少4位年份文件夹，例如2026：'+str(raw))
    m=re.fullmatch(r'(0?[1-9]|1[0-2])月?', month_name)
    d=re.fullmatch(r'(\d{1,2})[.\-_](\d{1,2})',day_name)
    if not m or not d:
        raise ValueError('月份目录应为10月、日期目录应为10.30；当前是：'+str(raw))
    month=int(m.group(1)); daymonth=int(d.group(1)); day=int(d.group(2))
    if month!=daymonth:
        raise ValueError(f'月份文件夹「{month_name}」与日期文件夹「{day_name}」不一致')
    dt=date(int(year_name), month, day)
    return category.parent,dt,raw


def main():
    parser=argparse.ArgumentParser(description='输入01日期文件夹路径，自动解析日期；按回车逐次执行，不后台监控')
    parser.add_argument('--orders',help='01分类下的年份/月份/日期订单文件夹')
    parser.add_argument('--once',action='store_true',help='只执行一轮（用于调试和历史对比）')
    parser.add_argument('--verbose',action='store_true',help='显示完整执行明细（默认精简终端；完整明细仍写入运行日志.txt）')
    args=parser.parse_args()
    global VERBOSE_OUTPUT
    VERBOSE_OUTPUT = args.verbose
    print('周二周五一体化发货助手 V4.12（云仓补运费清单在08；精简终端）')
    print('请粘贴8份订单所在的【01日期文件夹】完整路径，回车确认。')
    if args.orders:
        entered=args.orders
    else:
        try:
            entered=input('订单文件夹路径：')
        except EOFError:
            return 3
    try:
        root,dt,raw=order_path_settings(entered)
    except Exception as exc:
        msg('[路径错误] '+str(exc))
        return 3
    msg('自动识别发货日期：'+str(dt))
    msg('自动识别分类根目录：'+str(root))
    msg('本次01日期目录：'+str(raw))
    msg('八份原始订单推荐放在：'+str(raw/'订单文件'))
    msg('商品统计表直接生成到：'+str(raw/'商品统计表.xlsx'))
    msg('提示：盘符和总文件夹名称可以变化；程序根据01目录的上一级自动找到03/05/07/08/09。')
    msg('原则：按回车生成门店3份+云仓2份（补运费另增1份）；发货前审核F/G/H，改动后再按回车更新未发货的模板。')
    msg('门店模板：将模板.xlsx放到Python程序同目录，或分类根目录；保留全部列只写E/F/G/I（补运费另写P）。')
    while True:
        try:
            result=run(root,dt,str(raw))
        except Exception as exc:
            msg('\n[本轮发生错误] '+str(exc))
            if DETAIL_LOG_FILE is not None:
                with DETAIL_LOG_FILE.open('a',encoding='utf-8') as ef:
                    ef.write(traceback.format_exc()+'\n')
            if VERBOSE_OUTPUT:
                traceback.print_exc()
            else:
                print('  [错误详情] 可查看 01日期目录/合并结果/运行日志.txt',flush=True)
            result=3
        if args.once:
            return result
        try:
            response=input('\n按【回车】检查当前批次；粘贴另一个01日期目录可切换批次；输入 q 退出：').strip()
        except (EOFError,KeyboardInterrupt):
            msg('\n已退出程序。')
            return 0
        if response.lower() in ('q','quit','exit','退出'):
            msg('已退出程序。')
            return 0
        if response:
            # 粘贴新的01日期目录，或复制带引号的路径即可切换盘符和批次。
            # 非法路径仅提示，不执行上一个批次，以免误操作。
            try:
                new_root,new_dt,new_raw=order_path_settings(response)
            except Exception as exc:
                msg('[未切换批次] '+str(exc))
                msg('请直接粘贴【01分类/年份/月份/日期】目录，或只按回车检查当前批次。')
                continue
            root,dt,raw=new_root,new_dt,new_raw
            msg('\n[已切换批次] 日期：'+str(dt))
            msg('  新总目录：'+str(root))
            msg('  新订单目录：'+str(raw))


if __name__=='__main__':
    sys.exit(main())
