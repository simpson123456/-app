import streamlit as st
import tabula
import pandas as pd
import numpy as np
import re
import pdfplumber
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb

uploaded_file = st.file_uploader("选择PDF文件:", type="pdf")

if uploaded_file is not None:
    output_file_name = str(st.text_input('输出文件名称：')) + '.xlsx'
    input_file_name = str(uploaded_file.name) + '.pdf'
    pdf = pdfplumber.open(input_file_name)
    pdf_page_num = len(pdf.pages)

    sheet_nums = []  # 预录入编号
    export_num = []  # 海关编号
    export_port = []  # 出口口岸
    record_num = []  # 备案号
    export_date = []  # 出口日期
    declare_date = []  # 申报日期
    domestic_consignee = []  # 境内收发货人
    shipping_method = []  # 运输方式
    means_of_transport_name = []  # 运输工具名称
    bill_of_lading_number = []  # 提运单号
    prod_entity = []  # 生产销售单位
    trading_method = []  # 贸易方式
    nature_of_exemption = []  # 征免性质
    exchange_settlement = []  # 结汇方式
    license_number = []  # 许可证号
    country_of_arrival = []  # 运抵国
    transshipment_port = []  # 指运港
    domestic_source_of_goods = []  # 境内货源地
    approval_no = []  # 批准文号
    closing_method = []  # 成交方式
    freight = []  # 运费
    premium = []  # 保费
    miscellaneous = []  # 杂费
    contract_num = []  # 合同协议号
    num_of_pieces = []  # 件数
    type_of_packaging = []  # 包装种类
    gross_weight = []  # 毛重
    net_weight = []  # 净重
    container_no = []  # 集装箱号
    documents_attached = []  # 随附单证
    manufacturer = []  # 生产厂家
    shipping_marks_and_remarks = []  # 标记唛码及备注
    good_num = []  # 商品编号
    good_name = []  # 商品名称、规格
    good_quan = []  # 数量及单位
    country_of_final_destination = []  # 最终目的国
    price = []  # 单价
    sum_price = []  # 总价
    print_date = []  # 打印日期

    for page_num in range(pdf_page_num):
        if '主页' in pdf.pages[page_num].extract_text():
            sheet_nums_i = tabula.read_pdf(input_file_name,pages=page_num+1, area=[1.55*72, 1.52*72, 1.67*72, 2.77*72])[0].columns[0]
            export_num_i = tabula.read_pdf(input_file_name,pages=page_num+1, area=[1.55*72, 6.5*72, 1.67*72, 7.77*72])[0].columns[0]
            export_port_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[1.8*72, 0.575*72, 2.15*72, 2.745*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[1.8*72, 0.575*72, 2.15*72, 2.745*72])[0].iloc[0,0]
            record_num_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[1.8*72, 2.74*72, 2.15*72, 4.9*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[1.8*72, 2.74*72, 2.15*72, 4.9*72])[0].iloc[0,0]
            export_date_i = "".join(list(filter(None, sum(pdf.pages[page_num].extract_tables()[0], []))))[
                    "".join(list(filter(None, sum(pdf.pages[page_num].extract_tables()[0], [])))).find("出口日期\n") + len("出口日期\n"):"".join(list(filter(None, sum(pdf.pages[page_num].extract_tables()[0], [])))).find("申报日期")]
            declare_date_i = "".join(list(filter(None, sum(pdf.pages[page_num].extract_tables()[0], []))))[
                    "".join(list(filter(None, sum(pdf.pages[page_num].extract_tables()[0], [])))).find("申报日期\n") + len("申报日期\n"):"".join(list(filter(None, sum(pdf.pages[page_num].extract_tables()[0], [])))).find("境内收发货人")]
            domestic_consignee_i= '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.15*72, 0.575*72, 2.525*72, 2.745*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.15*72, 0.575*72, 2.525*72, 2.745*72])[0].iloc[0,0]
            shipping_method_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.15*72, 2.745*72, 2.525*72, 4.13*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.15*72, 2.745*72, 2.525*72, 4.13*72])[0].iloc[0,0]
            means_of_transport_name_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.15*72, 4.13*72, 2.525*72, 6.205*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.15*72, 4.13*72, 2.525*72, 6.205*72])[0].iloc[0,0]
            bill_of_lading_number_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.15*72, 6.205*72, 2.525*72, 7.765*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.15*72, 6.205*72, 2.525*72, 7.765*72])[0].iloc[0,0]
            prod_entity_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.525*72, 0.575*72, 2.88*72, 2.745*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.525*72, 0.575*72, 2.88*72, 2.745*72])[0].iloc[0,0]
            trading_method_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.525*72, 2.745*72, 2.88*72, 4.9*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.525*72, 2.745*72, 2.88*72, 4.9*72])[0].iloc[0,0]
            nature_of_exemption_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.525*72, 4.9*72, 2.88*72, 6.8*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.525*72, 4.9*72, 2.88*72, 6.8*72])[0].iloc[0,0]
            exchange_settlement_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.525*72, 4.9*72, 2.88*72, 7.765*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.525*72, 4.9*72, 2.88*72, 7.765*72])[0].iloc[0,0]
            license_number_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.88*72, 0.575*72, 3.235*72, 2.745*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.88*72, 0.575*72, 3.235*72, 2.745*72])[0].iloc[0,0]
            country_of_arrival_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.88*72, 2.745*72, 3.235*72, 4.9*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.88*72, 2.745*72, 3.235*72, 4.9*72])[0].iloc[0,0]
            transshipment_port_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.88*72, 4.9*72, 3.235*72, 6.205*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.88*72, 4.9*72, 3.235*72, 6.205*72])[0].iloc[0,0]
            domestic_source_of_goods_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.88*72, 6.205*72, 3.235*72, 7.765*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[2.88*72, 6.205*72, 3.235*72, 7.765*72])[0].iloc[0,0]
            approval_no_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.235*72, 0.575*72, 3.625*72, 2.745*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.235*72, 0.575*72, 3.625*72, 2.745*72])[0].iloc[0,0]
            closing_method_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.235*72, 2.745*72, 3.625*72, 4.13*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.235*72, 2.745*72, 3.625*72, 4.13*72])[0].iloc[0,0]
            freight_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.235*72, 4.13*72, 3.625*72, 5.365*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.235*72, 4.13 *72, 3.625*72, 5.365*72])[0].iloc[0,0]
            premium_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.235*72, 5.365*72, 3.625*72, 6.205*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.235*72, 5.365*72, 3.625*72, 6.205*72])[0].iloc[0,0]
            miscellaneous_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.235*72, 6.205*72, 3.625*72, 7.765*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.235*72, 6.205*72, 3.625*72, 7.765*72])[0].iloc[0,0]
            contract_num_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.625*72, 0.575*72, 3.975*72, 2.745*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.625*72, 0.575*72, 3.975*72, 2.745*72])[0].iloc[0,0]
            num_of_pieces_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.625*72, 2.745*72, 3.975*72, 4.13*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.625*72, 2.745*72, 3.975*72, 4.13*72])[0].iloc[0,0]
            type_of_packaging_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.625*72, 4.13*72, 3.975*72, 5.365*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.625*72, 4.13*72, 3.975*72, 5.365*72])[0].iloc[0,0]
            gross_weight_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.625*72, 5.365*72, 3.975*72, 6.8*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.625*72, 5.365*72, 3.975*72, 6.8*72])[0].iloc[0,0]
            net_weight_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.625*72, 6.8*72, 3.975*72, 7.765*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.625*72, 6.8*72, 3.975*72, 7.765*72])[0].iloc[0,0]
            container_no_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.975*72, 0.575*72, 4.325*72, 2.745*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.975*72, 0.575*72, 4.325*72, 2.745*72])[0].iloc[0,0]
            documents_attached_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.975*72, 2.745*72, 4.325*72, 6.205*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.975*72, 2.745*72, 4.325*72, 6.205*72])[0].iloc[0,0]
            manufacturer_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.975*72, 6.205*72, 4.325*72, 7.765*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[3.975*72, 6.205*72, 4.325*72, 7.765*72])[0].iloc[0,0]
            shipping_marks_and_remarks_i = '' if tabula.read_pdf(input_file_name,pages=page_num+1, area=[4.325*72, 0.575*72, 4.67*72, 7.765*72])[0].empty \
                else tabula.read_pdf(input_file_name,pages=page_num+1, area=[4.325*72, 0.575*72, 4.67*72, 7.765*72])[0].iloc[0,0]
            print_date_i = pdf.pages[page_num].extract_text()[pdf.pages[page_num].extract_text().find("打印日期：") + len("打印日期："):len(
                pdf.pages[page_num].extract_text())]

            for row in range(6):
                if tabula.read_pdf(input_file_name, pages=page_num + 1, area=[(5.027+row*0.458) * 72, 0.575 * 72, (5.485+row*0.458) * 72, 7.765 * 72])[0].empty is False:
                    good_num_i = tabula.read_pdf(input_file_name,pages=page_num+1, area=[(5.027+row*0.458) *72, 1.05*72, (5.485+row*0.458)*72, 1.725*72])[0].columns[0]
                    good_name_i = tabula.read_pdf(input_file_name,pages=page_num+1, area=[(5.027+row*0.458) *72, 1.725*72, (5.485+row*0.458)*72, 2.75*72])[0].columns[0]
                    good_quan_i = tabula.read_pdf(input_file_name, pages=page_num + 1,
                        area=[(5.027+row*0.458) *72, 2.75*72, (5.485+row*0.458)*72, 3.6*72])[0].columns[0]+' '+str(tabula.read_pdf(input_file_name,pages=page_num+1, area=[(5.027+row*0.458) *72, 2.75*72, (5.485+row*0.458)*72, 3.6*72])[0].iloc[0,0])
                    country_of_final_destination_i = tabula.read_pdf(input_file_name, pages=page_num + 1,
                        area=[(5.027+row*0.458) *72, 3.6*72, (5.485+row*0.458)*72, 4.85*72])[0].columns[0]+str(tabula.read_pdf(input_file_name,pages=page_num+1, area=[(5.027+row*0.458) *72, 3.6*72, (5.485+row*0.458)*72, 4.85*72])[0].iloc[0,0])
                    price_i = tabula.read_pdf(input_file_name,pages=page_num+1, area=[(5.027+row*0.458) *72, 4.85*72, (5.485+row*0.458)*72, 5.4*72])[0].columns[0]
                    sum_price_i = tabula.read_pdf(input_file_name,pages=page_num+1, area=[(5.027+row*0.458) *72, 5.4*72, (5.485+row*0.458)*72, 6.25*72])[0].columns[0]

                    sheet_nums.append(sheet_nums_i)
                    export_num.append(export_num_i)
                    export_port.append(export_port_i)
                    record_num.append(record_num_i)
                    export_date.append(export_date_i)
                    declare_date.append(declare_date_i)
                    domestic_consignee.append(domestic_consignee_i)
                    shipping_method.append(shipping_method_i)
                    means_of_transport_name.append(means_of_transport_name_i)
                    bill_of_lading_number.append(bill_of_lading_number_i)
                    prod_entity.append(prod_entity_i)
                    trading_method.append(trading_method_i)
                    nature_of_exemption.append(nature_of_exemption_i)
                    exchange_settlement.append(exchange_settlement_i)
                    license_number.append(license_number_i)
                    country_of_arrival.append(country_of_arrival_i)
                    transshipment_port.append(transshipment_port_i)
                    domestic_source_of_goods.append(domestic_source_of_goods_i)
                    approval_no.append(approval_no_i)
                    closing_method.append(closing_method_i)
                    freight.append(freight_i)
                    premium.append(premium_i)
                    miscellaneous.append(miscellaneous_i)
                    contract_num.append(contract_num_i)
                    num_of_pieces.append(num_of_pieces_i)
                    type_of_packaging.append(type_of_packaging_i)
                    gross_weight.append(gross_weight_i)
                    net_weight.append(net_weight_i)
                    container_no.append(container_no_i)
                    documents_attached.append(documents_attached_i)
                    manufacturer.append(manufacturer_i)
                    shipping_marks_and_remarks.append(shipping_marks_and_remarks_i)
                    good_num.append(good_num_i)
                    good_name.append(good_name_i)
                    good_quan.append(good_quan_i)
                    country_of_final_destination.append(country_of_final_destination_i)
                    price.append(price_i)
                    sum_price.append(sum_price_i)
                    print_date.append(print_date_i)
                else:
                    break
        else:
            sheet_nums_i = \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[1.55 * 72-0.122*72, 1.52 * 72, 1.67 * 72-0.122*72, 2.77 * 72])[
                0].columns[0]
            export_num_i = \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[1.55 * 72-0.122*72, 6.5 * 72, 1.67 * 72-0.122*72, 7.77 * 72])[
                0].columns[0]
            export_port_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[1.8 * 72-0.122*72, 0.575 * 72, 2.15 * 72-0.122*72, 2.745 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[1.8 * 72-0.122*72, 0.575 * 72, 2.15 * 72-0.122*72, 2.745 * 72])[
                0].iloc[0, 0]
            record_num_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[1.8 * 72-0.122*72, 2.74 * 72, 2.15 * 72-0.122*72, 4.9 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[1.8 * 72-0.122*72, 2.74 * 72, 2.15 * 72-0.122*72, 4.9 * 72])[
                0].iloc[0, 0]
            export_date_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[1.8 * 72-0.122*72, 4.9 * 72, 2.15 * 72-0.122*72, 6.205 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[1.8 * 72-0.122*72, 4.9 * 72, 2.15 * 72-0.122*72, 6.205 * 72])[
                0].iloc[0, 0]
            declare_date_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[1.8 * 72-0.122*72, 6.205 * 72, 2.15 * 72-0.122*72, 7.765 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[1.8 * 72-0.122*72, 6.205 * 72, 2.15 * 72-0.122*72, 7.765 * 72])[
                0].iloc[0, 0]
            domestic_consignee_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.15 * 72-0.122*72, 0.575 * 72, 2.525 * 72-0.122*72, 2.745 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.15 * 72-0.122*72, 0.575 * 72, 2.525 * 72-0.122*72, 2.745 * 72])[
                0].iloc[0, 0]
            shipping_method_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.15 * 72-0.122*72, 2.745 * 72, 2.525 * 72-0.122*72, 4.13 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.15 * 72-0.122*72, 2.745 * 72, 2.525 * 72-0.122*72, 4.13 * 72])[
                0].iloc[0, 0]
            means_of_transport_name_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.15 * 72-0.122*72, 4.13 * 72, 2.525 * 72-0.122*72, 6.205 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.15 * 72-0.122*72, 4.13 * 72, 2.525 * 72-0.122*72, 6.205 * 72])[
                0].iloc[0, 0]
            bill_of_lading_number_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.15 * 72-0.122*72, 6.205 * 72, 2.525 * 72-0.122*72, 7.765 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.15 * 72-0.122*72, 6.205 * 72, 2.525 * 72-0.122*72, 7.765 * 72])[
                0].iloc[0, 0]
            prod_entity_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.525 * 72-0.122*72, 0.575 * 72, 2.88 * 72-0.122*72, 2.745 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.525 * 72-0.122*72, 0.575 * 72, 2.88 * 72-0.122*72, 2.745 * 72])[
                0].iloc[0, 0]
            trading_method_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.525 * 72-0.122*72, 2.745 * 72, 2.88 * 72-0.122*72, 4.9 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.525 * 72-0.122*72, 2.745 * 72, 2.88 * 72-0.122*72, 4.9 * 72])[
                0].iloc[0, 0]
            nature_of_exemption_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.525 * 72-0.122*72, 4.9 * 72, 2.88 * 72-0.122*72, 6.8 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.525 * 72-0.122*72, 4.9 * 72, 2.88 * 72-0.122*72, 6.8 * 72])[
                0].iloc[0, 0]
            exchange_settlement_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.525 * 72-0.122*72, 4.9 * 72, 2.88 * 72-0.122*72, 7.765 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.525 * 72-0.122*72, 4.9 * 72, 2.88 * 72-0.122*72, 7.765 * 72])[
                0].iloc[0, 0]
            license_number_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.88 * 72-0.122*72, 0.575 * 72, 3.235 * 72-0.122*72, 2.745 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.88 * 72-0.122*72, 0.575 * 72, 3.235 * 72-0.122*72, 2.745 * 72])[
                0].iloc[0, 0]
            country_of_arrival_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.88 * 72-0.122*72, 2.745 * 72, 3.235 * 72-0.122*72, 4.9 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.88 * 72-0.122*72, 2.745 * 72, 3.235 * 72-0.122*72, 4.9 * 72])[
                0].iloc[0, 0]
            transshipment_port_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.88 * 72-0.122*72, 4.9 * 72, 3.235 * 72-0.122*72, 6.205 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.88 * 72-0.122*72, 4.9 * 72, 3.235 * 72-0.122*72, 6.205 * 72])[
                0].iloc[0, 0]
            domestic_source_of_goods_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.88 * 72-0.122*72, 6.205 * 72, 3.235 * 72-0.122*72, 7.765 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[2.88 * 72-0.122*72, 6.205 * 72, 3.235 * 72-0.122*72, 7.765 * 72])[
                0].iloc[0, 0]
            approval_no_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.235 * 72-0.122*72, 0.575 * 72, 3.625 * 72-0.122*72, 2.745 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.235 * 72-0.122*72, 0.575 * 72, 3.625 * 72-0.122*72, 2.745 * 72])[
                0].iloc[0, 0]
            closing_method_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.235 * 72-0.122*72, 2.745 * 72, 3.625 * 72-0.122*72, 4.13 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.235 * 72-0.122*72, 2.745 * 72, 3.625 * 72-0.122*72, 4.13 * 72])[
                0].iloc[0, 0]
            freight_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.235 * 72-0.122*72, 4.13 * 72, 3.625 * 72-0.122*72, 5.365 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.235 * 72-0.122*72, 4.13 * 72, 3.625 * 72-0.122*72, 5.365 * 72])[
                0].iloc[0, 0]
            premium_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.235 * 72-0.122*72, 5.365 * 72, 3.625 * 72-0.122*72, 6.205 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.235 * 72-0.122*72, 5.365 * 72, 3.625 * 72-0.122*72, 6.205 * 72])[
                0].iloc[0, 0]
            miscellaneous_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.235 * 72-0.122*72, 6.205 * 72, 3.625 * 72-0.122*72, 7.765 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.235 * 72-0.122*72, 6.205 * 72, 3.625 * 72-0.122*72, 7.765 * 72])[
                0].iloc[0, 0]
            contract_num_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.625 * 72-0.122*72, 0.575 * 72, 3.975 * 72-0.122*72, 2.745 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.625 * 72-0.122*72, 0.575 * 72, 3.975 * 72-0.122*72, 2.745 * 72])[
                0].iloc[0, 0]
            num_of_pieces_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.625 * 72-0.122*72, 2.745 * 72, 3.975 * 72-0.122*72, 4.13 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.625 * 72-0.122*72, 2.745 * 72, 3.975 * 72-0.122*72, 4.13 * 72])[
                0].iloc[0, 0]
            type_of_packaging_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.625 * 72-0.122*72, 4.13 * 72, 3.975 * 72-0.122*72, 5.365 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.625 * 72-0.122*72, 4.13 * 72, 3.975 * 72-0.122*72, 5.365 * 72])[
                0].iloc[0, 0]
            gross_weight_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.625 * 72-0.122*72, 5.365 * 72, 3.975 * 72-0.122*72, 6.8 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.625 * 72-0.122*72, 5.365 * 72, 3.975 * 72-0.122*72, 6.8 * 72])[
                0].iloc[0, 0]
            net_weight_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.625 * 72-0.122*72, 6.8 * 72, 3.975 * 72-0.122*72, 7.765 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.625 * 72-0.122*72, 6.8 * 72, 3.975 * 72-0.122*72, 7.765 * 72])[
                0].iloc[0, 0]
            container_no_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.975 * 72-0.122*72, 0.575 * 72, 4.325 * 72-0.122*72, 2.745 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.975 * 72-0.122*72, 0.575 * 72, 4.325 * 72-0.122*72, 2.745 * 72])[
                0].iloc[0, 0]
            documents_attached_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.975 * 72-0.122*72, 2.745 * 72, 4.325 * 72-0.122*72, 6.205 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.975 * 72-0.122*72, 2.745 * 72, 4.325 * 72-0.122*72, 6.205 * 72])[
                0].iloc[0, 0]
            manufacturer_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.975 * 72-0.122*72, 6.205 * 72, 4.325 * 72-0.122*72, 7.765 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[3.975 * 72-0.122*72, 6.205 * 72, 4.325 * 72-0.122*72, 7.765 * 72])[
                0].iloc[0, 0]
            shipping_marks_and_remarks_i = '' if \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[4.325 * 72-0.122*72, 0.575 * 72, 4.67 * 72-0.122*72, 7.765 * 72])[
                0].empty \
                else \
            tabula.read_pdf(input_file_name, pages=page_num + 1, area=[4.325 * 72-0.122*72, 0.575 * 72, 4.67 * 72-0.122*72, 7.765 * 72])[
                0].iloc[0, 0]
            print_date_i = pdf.pages[page_num].extract_text()[
                           pdf.pages[page_num].extract_text().find("打印日期：") + len("打印日期："):len(
                               pdf.pages[page_num].extract_text())]

            for row in range(6):
                if tabula.read_pdf(input_file_name, pages=page_num + 1, area=[(5.027+row*0.458) * 72-0.122*72, 0.575 * 72, (5.485+row*0.458) * 72-0.122*72, 7.765 * 72])[0].empty is False:
                    good_num_i = tabula.read_pdf(input_file_name, pages=page_num + 1,
                                                 area=[(5.027 + row * 0.458) * 72-0.122*72, 1.05 * 72,
                                                       (5.485 + row * 0.458) * 72-0.122*72, 1.725 * 72])[0].columns[0]
                    good_name_i = tabula.read_pdf(input_file_name, pages=page_num + 1,
                                                  area=[(5.027 + row * 0.458) * 72-0.122*72, 1.725 * 72,
                                                        (5.485 + row * 0.458) * 72-0.122*72, 2.75 * 72])[0].columns[0]
                    good_quan_i = tabula.read_pdf(input_file_name, pages=page_num + 1,
                                                  area=[(5.027 + row * 0.458) * 72-0.122*72, 2.75 * 72,
                                                        (5.485 + row * 0.458) * 72-0.122*72, 3.6 * 72])[0].columns[
                                      0] + ' ' + str(tabula.read_pdf(input_file_name, pages=page_num + 1,
                                                                     area=[(5.027 + row * 0.458) * 72-0.122*72, 2.75 * 72,
                                                                           (5.485 + row * 0.458) * 72-0.122*72, 3.6 * 72])[
                                                         0].iloc[0, 0])
                    country_of_final_destination_i = tabula.read_pdf(input_file_name, pages=page_num + 1,
                                                                     area=[(5.027 + row * 0.458) * 72-0.122*72, 3.6 * 72,
                                                                           (5.485 + row * 0.458) * 72-0.122*72, 4.85 * 72])[
                                                         0].columns[0] + str(
                        tabula.read_pdf(input_file_name, pages=page_num + 1,
                                        area=[(5.027 + row * 0.458) * 72-0.122*72, 3.6 * 72, (5.485 + row * 0.458) * 72-0.122*72,
                                              4.85 * 72])[0].iloc[0, 0])
                    price_i = tabula.read_pdf(input_file_name, pages=page_num + 1,
                                              area=[(5.027 + row * 0.458) * 72-0.122*72, 4.85 * 72, (5.485 + row * 0.458) * 72-0.122*72,
                                                    5.4 * 72])[0].columns[0]
                    sum_price_i = tabula.read_pdf(input_file_name, pages=page_num + 1,
                                                  area=[(5.027 + row * 0.458) * 72-0.122*72, 5.4 * 72,
                                                        (5.485 + row * 0.458) * 72-0.122*72, 6.25 * 72])[0].columns[0]

                    sheet_nums.append(sheet_nums_i)
                    export_num.append(export_num_i)
                    export_port.append(export_port_i)
                    record_num.append(record_num_i)
                    export_date.append(export_date_i)
                    declare_date.append(declare_date_i)
                    domestic_consignee.append(domestic_consignee_i)
                    shipping_method.append(shipping_method_i)
                    means_of_transport_name.append(means_of_transport_name_i)
                    bill_of_lading_number.append(bill_of_lading_number_i)
                    prod_entity.append(prod_entity_i)
                    trading_method.append(trading_method_i)
                    nature_of_exemption.append(nature_of_exemption_i)
                    exchange_settlement.append(exchange_settlement_i)
                    license_number.append(license_number_i)
                    country_of_arrival.append(country_of_arrival_i)
                    transshipment_port.append(transshipment_port_i)
                    domestic_source_of_goods.append(domestic_source_of_goods_i)
                    approval_no.append(approval_no_i)
                    closing_method.append(closing_method_i)
                    freight.append(freight_i)
                    premium.append(premium_i)
                    miscellaneous.append(miscellaneous_i)
                    contract_num.append(contract_num_i)
                    num_of_pieces.append(num_of_pieces_i)
                    type_of_packaging.append(type_of_packaging_i)
                    gross_weight.append(gross_weight_i)
                    net_weight.append(net_weight_i)
                    container_no.append(container_no_i)
                    documents_attached.append(documents_attached_i)
                    manufacturer.append(manufacturer_i)
                    shipping_marks_and_remarks.append(shipping_marks_and_remarks_i)
                    good_num.append(good_num_i)
                    good_name.append(good_name_i)
                    good_quan.append(good_quan_i)
                    country_of_final_destination.append(country_of_final_destination_i)
                    price.append(price_i)
                    sum_price.append(sum_price_i)
                    print_date.append(print_date_i)
                else:
                    break

    data = {
        "预录入编号": sheet_nums,
        "海关编号": export_num,
        "出口口岸(0101)": export_port,
        "备案号": record_num,
        "出口日期": export_date,
        "申报日期": declare_date,
        "境内收发货人(1211940259)": domestic_consignee,
        "运输方式": shipping_method,
        "运输工具名称": means_of_transport_name,
        "提运单号": bill_of_lading_number,
        "生产销售单位": prod_entity,
        "贸易方式(0110)": trading_method,
        "征免性质": nature_of_exemption,
        "结汇方式": exchange_settlement,
        "许可证号": license_number,
        "运抵国（地区）(419)": country_of_arrival,
        "指运港": transshipment_port,
        "境内货源地(12119)": domestic_source_of_goods,
        "批准文号": approval_no,
        "成交方式": closing_method,
        "运费": freight,
        "保费": premium,
        "杂费": miscellaneous,
        "合同协议号": contract_num,
        "件数": num_of_pieces,
        "包装种类": type_of_packaging,
        "毛重（千克）": gross_weight,
        "净重（千克）": net_weight,
        "集装箱号": container_no,
        "随附单证": documents_attached,
        "生产厂家": manufacturer,
        "标记唛码及备注": shipping_marks_and_remarks,
        "商品编号": good_num,
        "商品名称、规格型号": good_name,
        "数量及单位": good_quan,
        "最终目的国（地区）": country_of_final_destination,
        "单价": price,
        "总价": sum_price,
        "打印日期": print_date
    }

    data = pd.DataFrame.from_dict(data)

    def to_excel(df):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        writer._save()  # writer.save() 版本问题使用：_save()
        processed_data = output.getvalue()
        return processed_data

    df_xlsx = to_excel(data)

    st.download_button(
        label = "📥下载文件至本地",
        data = df_xlsx,
        file_name = output_file_name,
    )
