import streamlit as st
import pandas as pd
import re
import pdfplumber
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb

uploaded_file = st.file_uploader("选择PDF文件:", type="pdf")
if uploaded_file is not None:
    a = st.text_input('输出文件名称：')
    st.write(uploaded_file.name)

    b = uploaded_file

    pdf = pdfplumber.open(b)

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

    for page in pdf.pages:
        contents_all = page.extract_text()
        contents_table = "".join(list(filter(None, sum(page.extract_tables()[0], []))))
        sheet_nums_i = contents_all[
                       contents_all.find("预录入编号：") + len("预录入编号："):contents_all.find(" 海关编号")]
        export_num_i = contents_all[
                       contents_all.find("海关编号：") + len("海关编号："):contents_all.find("\n出口口岸")]
        export_port_i = contents_table[
                        contents_table.find("出口口岸") + len("出口口岸(0101)\n"):contents_table.find("备案号")]
        record_num_i = contents_table[contents_table.find("备案号") + len("备案号"):contents_table.find("出口日期")]
        export_date_i = contents_table[
                        contents_table.find("出口日期\n") + len("出口日期\n"):contents_table.find("申报日期")]
        declare_date_i = contents_table[
                         contents_table.find("申报日期\n") + len("申报日期\n"):contents_table.find("境内收发货人")]
        domestic_consignee_i = contents_table[contents_table.find("境内收发货人") + len(
            "境内收发货人(1211940259)\n"):contents_table.find("运输方式")]
        shipping_method_i = contents_table[
                            contents_table.find("运输方式\n") + len("运输方式\n"):contents_table.find(
                                "运输工具名称")]
        means_of_transport_name_i = contents_table[
                                    contents_table.find("运输工具名称") + len("运输工具名称"):contents_table.find(
                                        "提运单号")]
        bill_of_lading_number_i = contents_table[
                                  contents_table.find("提运单号") + len("提运单号"):contents_table.find(
                                      "生产销售单位")]
        prod_entity_i = contents_table[
                        contents_table.find("生产销售单位\n") + len("生产销售单位\n"):contents_table.find(
                            "贸易方式")]
        trading_method_i = contents_table[
                           contents_table.find("贸易方式") + len("贸易方式(0110)\n"):contents_table.find(
                               "征免性质")]
        nature_of_exemption_i = contents_table[
                                contents_table.find("征免性质") + len("征免性质"):contents_table.find("结汇方式")]
        exchange_settlement_i = contents_table[
                                contents_table.find("结汇方式") + len("结汇方式"):contents_table.find("许可证号")]
        license_number_i = contents_table[
                           contents_table.find("许可证号") + len("许可证号"):contents_table.find("运抵国（地区）")]
        country_of_arrival_i = contents_table[
                               contents_table.find("运抵国（地区）") + len("运抵国（地区）(419)\n"):contents_table.find(
                                   "指运港")]
        transshipment_port_i = contents_table[
                               contents_table.find("指运港") + len("指运港"):contents_table.find("境内货源地")]
        domestic_source_of_goods_i = contents_table[
                                     contents_table.find("境内货源地") + len(
                                         "境内货源地(12119)\n"):contents_table.find(
                                         "批准文号")]
        approval_no_i = contents_table[
                        contents_table.find("批准文号") + len("批准文号"):contents_table.find("成交方式")]
        closing_method_i = contents_table[
                           contents_table.find("成交方式\n") + len("成交方式\n"):contents_table.find("运费")]
        freight_i = contents_table[contents_table.find("运费\n") + len("运费\n"):contents_table.find("保费\n")]
        premium_i = contents_table[contents_table.find("保费\n") + len("保费\n"):contents_table.find("杂费\n")]
        miscellaneous_i = contents_table[
                          contents_table.find("杂费\n") + len("杂费\n"):contents_table.find("合同协议号\n")]
        contract_num_i = contents_table[
                         contents_table.find("合同协议号\n") + len("合同协议号\n"):contents_table.find("件数")]
        num_of_pieces_i = contents_table[contents_table.find("件数") + len("件数"):contents_table.find("包装种类")]
        type_of_packaging_i = contents_table[
                              contents_table.find("包装种类") + len("包装种类"):contents_table.find("毛重（千克）")]
        gross_weight_i = contents_table[
                         contents_table.find("毛重（千克）") + len("毛重（千克）"):contents_table.find("净重（千克）")]
        net_weight_i = contents_table[
                       contents_table.find("净重（千克）") + len("净重（千克）"):contents_table.find("集装箱号")]
        container_no_i = contents_table[
                         contents_table.find("集装箱号") + len("集装箱号"):contents_table.find("随附单证")]
        documents_attached_i = contents_table[
                               contents_table.find("随附单证") + len("随附单证"):contents_table.find("生产厂家")]
        manufacturer_i = contents_table[
                         contents_table.find("生产厂家") + len("生产厂家"):contents_table.find("标记唛码及备注")]
        shipping_marks_and_remarks_i = contents_table[contents_table.find("标记唛码及备注") + len(
            "标记唛码及备注"):contents_table.find("商品名称、规格")]
        print_date_i = contents_all[contents_all.find("打印日期：") + len("打印日期："):len(contents_all)]

        for i in range(len(page.extract_tables()[0])):
            if "美元" in "".join(list(filter(None, page.extract_tables()[0][i]))):
                string = list(filter(None, page.extract_tables()[0][i]))[0]
                space_1 = string.find(" ")
                space_2 = string.find(" ", space_1 + 1)
                space_3 = string.find(" ", space_2 + 1)
                space_4 = string.find(" ", space_3 + 1)
                space_5 = string.find(" ", space_4 + 1)
                space_6 = string.find(" ", space_5 + 1)
                space_7 = string.find(" ", space_6 + 1)
                space_8 = string.find(" ", space_7 + 1)
                space_9 = string.find(" ", space_8 + 1)
                space_10 = string.find(" ", space_9 + 1)
                good_num_i = string[space_1 + 1:space_2]
                good_name_i = string[space_2 + 1:space_3]
                good_quan_i = string[space_3 + 1:space_4]
                country_of_final_destination_i = string[space_4 + 1:space_5]
                price_i = string[space_5 + 1:space_6]
                sum_price_i = string[space_6 + 1:space_7]

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
    file_name = a,
    mime = "text/csv",
)


