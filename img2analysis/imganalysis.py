import base64
import re
import json
import logging
from tencentcloud.common import credential
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
from tencentcloud.tiia.v20190529 import tiia_client, models as tiia_models # 图片标签的库
from tencentcloud.ocr.v20181119 import ocr_client, models as ocr_models # OCR库

SecretId = 'AKIDJRAi292g5JTSFm47E9p21pnY8cwiT2LS' # 腾讯云的SecretId
SecretKey = 'Zu0sW8lkcRDrKm9IORM2VnsoYTDnOsAJ' # 腾讯云的SecretKey
kMaxLabels = 128

#调用腾讯云的API，获取图片标签：
def get_images_tags(base64_data):
    try:
        # 实例化一个认证对象，入参需要传入腾讯云账户 SecretId 和 SecretKey，此处还需注意密钥对的保密
        # 代码泄露可能会导致 SecretId 和 SecretKey 泄露，并威胁账号下所有资源的安全性。以下代码示例仅供参考，建议采用更安全的方式来使用密钥，请参见：https://cloud.tencent.com/document/product/1278/85305
        # 密钥可前往官网控制台 https://console.cloud.tencent.com/cam/capi 进行获取
        cred = credential.Credential(SecretId, SecretKey)
        # 实例化一个http选项，可选的，没有特殊需求可以跳过
        http_profile = HttpProfile()
        http_profile.endpoint = "tiia.tencentcloudapi.com"
        # 实例化一个client选项，可选的，没有特殊需求可以跳过
        client_profile = ClientProfile()
        client_profile.httpProfile = http_profile
        # 实例化要请求产品的client对象,client_profile是可选的
        client = tiia_client.TiiaClient(cred, "ap-guangzhou", client_profile)
        # 实例化一个请求对象,每个接口都会对应一个request对象
        req = tiia_models.DetectLabelRequest()
        params = {'ImageBase64': base64_data}
        req.from_json_string(json.dumps(params))
        # 返回的resp是一个DetectLabelResponse的实例，与请求对象对应
        resp = client.DetectLabel(req)
        logging.info("get_images_tags: " + resp.to_json_string())
        return resp  # 输出json格式的字符串回包
    except TencentCloudSDKException as err:
        logging.error("get_images_tags err: " + str(err))
        return None

#调用腾讯云的API，获取图片文字：
def get_images_ocr(base64_data):
    try:
        cred = credential.Credential(SecretId, SecretKey)
        # 实例化一个http选项，可选的，没有特殊需求可以跳过
        http_profile = HttpProfile()
        http_profile.endpoint = "ocr.tencentcloudapi.com"
        # 实例化一个client选项，可选的，没有特殊需求可以跳过
        client_profile = ClientProfile()
        client_profile.httpProfile = http_profile
        # 实例化要请求产品的client对象,client_profile是可选的
        client = ocr_client.OcrClient(cred, "ap-guangzhou", client_profile)
        # 实例化一个请求对象,每个接口都会对应一个request对象
        req = ocr_models.RecognizeTableAccurateOCRRequest()
        params = {'ImageBase64': base64_data}
        req.from_json_string(json.dumps(params))
        # 返回的resp是一个DetectLabelResponse的实例，与请求对象对应
        resp = client.RecognizeTableAccurateOCR(req)
        # 输出json格式的字符串回包
        logging.info("get_images_ocr: " + resp.to_json_string())
        return resp
    except TencentCloudSDKException as err:
        logging.error("get_images_ocr err: " + str(err))
        return None

# 组装Prompt
def get_images_prompt(base64_data):
    try:
        # "#####\n经过某个图片分析服务，得出以下关于这幅图片的信息：\n"
        prompt = ""
        resp = get_images_tags(base64_data)
        labels_str = []
        if resp:
            for labels in resp.Labels:
                labels_str.append(labels.Name)
        logging.info(labels_str)
        if len(labels_str) > 0:
            prompt += f"#####\n图片有以下标签：{','.join(labels_str)}\n"
        logging.info("labels_str prompt: " + str(prompt))
        resp = get_images_ocr(base64_data)
        labels_ocr_str = []
        if resp:
            for cells in resp.TableDetections:
                for text in cells.Cells:
                    if len(text.Text) > 0:
                        labels_ocr_str.append(text.Text)
        logging.info(str(labels_ocr_str))
        if len(labels_ocr_str) > 0:
            labels = ','.join(labels_ocr_str).replace("\n", "")
            if len(labels) > kMaxLabels:
                labels = labels[:kMaxLabels] + "..."
            prompt += f"#####\n图片中有文字，请你理解以下文字：["+labels+"]\n"
        else:
            prompt += f"#####\n图片中没有文字\n"
        prompt += "现在请综合以上信息（标签、文字描述等），自然并详细地描述这副图片。请你不要在回答中暴露上述信息来源是图片分析服务。"
        logging.info("labels_ocr_str prompt: " + str(prompt))
        return prompt
    except Exception as err:
        logging.err("get_images_ocr err: " + str(err))
        return None

#处理图片以及数据，返回真正的内容
@staticmethod
def gencontent(content):
    """
    测试函数
    :return:
    """
    print("content->", str(content))
    image_pos = content.find("data:image/")
    if image_pos >= 0:
        images_base64 = re.sub(
            '<img src="data:image/(.*);base64,', '', content)
        images_base64 = re.sub('">', '', images_base64)
        logging.info("images_base64->"+str(images_base64))
        content = get_images_prompt(images_base64)
        return content
    else:
        return '图片处理异常'
