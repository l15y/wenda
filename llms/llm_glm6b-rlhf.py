from plugins.common import settings

def chat_init(history):
    history_formatted = None
    if history is not None:
        history_formatted = []
        tmp = []
        for i, old_chat in enumerate(history):
            if len(tmp) == 0 and old_chat['role'] == "user":
                tmp.append(old_chat['content'])
            elif old_chat['role'] == "AI" or old_chat['role'] == 'assistant':
                tmp.append(old_chat['content'])
                history_formatted.append(tuple(tmp))
                tmp = []
            else:
                continue
    return history_formatted


def chat_one(prompt, history_formatted, max_length, top_p, temperature, data):
    yield str(len(prompt))+'字正在计算'
    for response, history in model.stream_chat(tokenizer, prompt, history_formatted,
                                               max_length=max_length, top_p=top_p, temperature=temperature):
        yield response

def sum_values(dict):
    total = 0
    for value in dict.values():
        total += value
    return total

def dict_to_list(d):
    l = []
    for k, v in d.items():
        l.extend([k] * v)
    return l

def load_model():
    global model, tokenizer
    from transformers import AutoModel, AutoTokenizer

    num_trans_layers = 28
    strategy = ('->'.join([x.strip() for x in settings.llm.strategy.split('->')])).replace('->', ' -> ')
    s = [x.strip().split(' ') for x in strategy.split('->')]
    print(s)
    if len(s)>1:
        from accelerate import dispatch_model
        start_device = int(s[0][0].split(':')[1])
        #根据路径名判断，如果是glm2则使用专用devicemap，参见https://github.com/THUDM/ChatGLM2-6B/blob/main/utils.py Line23
        if "chatglm2" in settings.llm.path.lower():
            device_map = {'transformer.embedding.word_embeddings': 0,
            'transformer.encoder.final_layernorm': 0,
            'transformer.output_layer': 0,
            'transformer.rotary_pos_emb': 0,
            'lm_head': 0}
        else:
            device_map = {'transformer.word_embeddings': start_device,
            'transformer.final_layernorm': start_device, 'lm_head': start_device}
        
        n = {}
        for i in range(len(s)):
            si = s[i]
            if len(s[i]) > 2:
                ss = si[2]
                if ss.startswith('*'):
                        n[int(si[0].split(':')[1])]=int(ss[1:])
            else:
                n[int(si[0].split(':')[1])] = num_trans_layers+2-sum_values(n)
        n[start_device] -= 2
        n = dict_to_list(n)
        for i in range(num_trans_layers):
            #根据路径名判断，如果是glm2则使用专用devicemap，参见https://github.com/THUDM/ChatGLM2-6B/blob/main/utils.py Line23
            if "chatglm2" in settings.llm.path.lower():
                device_map[f'transformer.encoder.layers.{i}'] = n[i]
            else:
                device_map[f'transformer.layers.{i}'] = n[i]

    device, precision = s[0][0], s[0][1]
    
    tokenizer = AutoTokenizer.from_pretrained(
        settings.llm.path, local_files_only=True, trust_remote_code=True,revision="v1.1.0")
    model = AutoModel.from_pretrained(
        settings.llm.path, local_files_only=True, trust_remote_code=True, revision="v1.1.0")
    if not (settings.llm.lora == '' or settings.llm.lora == None):
        print('Lora模型地址', settings.llm.lora)
        from peft import PeftModel
        model = PeftModel.from_pretrained(model, settings.llm.lora,adapter_name=settings.llm.lora)
        
    # 根据设备执行不同的操作
    if device == 'cpu':
        # 如果是cpu，不做任何操作
        pass
    elif device == 'cuda':
        # 如果是gpu，把模型移动到显卡
        import torch
        if "chatglm2" in settings.llm.path and "int4" in settings.llm.path:
            model = model.cuda()
        elif not (precision.startswith('fp16i') and torch.cuda.get_device_properties(0).total_memory < 1.4e+10):
            model = model.cuda()
    elif len(s)>1 and device.startswith('cuda:'):
        pass
    else:
        # 如果是其他设备，报错并退出程序
        print('Error: 不受支持的设备')
        exit()
    # 根据精度执行不同的操作
    if precision == 'fp16':
        # 如果是fp16，把模型转化为半精度
        model = model.half()
    elif precision == 'fp32':
        # 如果是fp32，把模型转化为全精度
        model = model.float()
    elif precision.startswith('fp16i'):
        # 如果是fp16i开头，把模型转化为指定的精度
        # 从字符串中提取精度的数字部分
        bits = int(precision[5:])
        # 调用quantize方法，传入精度参数
        model = model.half().quantize(bits)
        if device == 'cuda':
            model = model.cuda()
        # model = model.half()
    elif precision.startswith('fp32i'):
        # 如果是fp32i开头，把模型转化为指定的精度
        # 从字符串中提取精度的数字部分
        bits = int(precision[5:])
        # 调用quantize方法，传入精度参数
        model = model.quantize(bits)
        if device == 'cuda':
            model = model.cuda()
        model = model.float()
    else:
        # 如果是其他精度，报错并退出程序
        print('Error: 不受支持的精度')
        exit()
    if len(s)>1:
        model = dispatch_model(model, device_map=device_map)
    model = model.eval()


if not (settings.llm.lora == '' or settings.llm.lora == None):
    from bottle import route, response, request
    @route('/lora_load_adapter', method=("POST","OPTIONS"))
    def load_adapter():
        # allowCROS()
        try:
            data = request.json
            lora_path=data.get("lora_path")
            adapter_name=data.get("adapter_name")
            model.load_adapter(lora_path, adapter_name=adapter_name)
            return "保存成功"
        except Exception as e:
            return str(e)
    @route('/lora_set_adapter', method=("POST","OPTIONS"))
    def set_adapter():
        # allowCROS()
        try:
            data = request.json
            adapter_name=data.get("adapter_name")
            model.set_adapter(adapter_name)
            return "保存成功"
        except Exception as e:
            return str(e)
