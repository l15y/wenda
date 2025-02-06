re_s1_p = {
    'chinese': '现在你需要帮助我完成关系抽取任务，给定的句子为："{}"\n\n给定关系列表：["所属行业","行业上级","上游原材料","下游产品","公司主营","产品小类"]\n\n在这个句子中，可能包含了哪些关系？\n请给出关系列表中的关系。\n如果不存在则回答：无\n按照元组形式回复，如 (关系1, 关系2, ……)：',
    'english': 'The given sentence is "{}"\n\nList of given relations: {}\n\nWhat relations in the given list might be included in this given sentence?\nIf not present, answer: none.\nRespond as a tuple, e.g. (relation 1, relation 2, ......):',
}

re_s2_p = {
    'chinese': '根据给定的句子："{}"\n\n，涵盖实体的类型分别为（{}，{}）且之间的关系为{}，给定的句子包含的实体有"{}"。如果有多组，则按组全部列出。\n如果不存在则回答：无\n按照表格形式回复，表格有两列且表头为（{}，{}）：',
    'english': 'According to the given sentence, the two entities are of type ("{}", "{}") and the relation between them is "{}", find the two entities and list them all by group if there are multiple groups.\nIf not present, answer: none.\nRespond in the form of a table with two columns and a header of ("{}", "{}"):',
}

re_s3_p = {
    'chinese': '{}\\n\\n提取上述句子中{}类型的实体，实体和类型要尽最大可能贴切，不要胡编乱造，并按照JSON格式输出，上述句子中不存在的信息用[\'原文中未提及\']来表示，多个值之间用\',\'分隔：',
    'english': 'According to the given sentence, the two entities are of type ("{}", "{}") and the relation between them is "{}", find the two entities and list them all by group if there are multiple groups.\nIf not present, answer: none.\nRespond in the form of a table with two columns and a header of ("{}", "{}"):',
}

df_ret = {
    'english': {
            'location-located_in': ['location', 'location'],
            'administrative_division-country': ['location', 'country'],
            'person-place_lived': ['person', 'location'],
            'person-company': ['person', 'organization'],
            'person-nationality': ['person', 'country'],
            'company-founders': ['organization', 'person'],
            'country-administrative_divisions': ['country', 'location'],
            'person-children': ['person', 'person'],
            'country-capital': ['country', 'city'],
            'deceased_person-place_of_death': ['person', 'location'],
            'neighborhood-neighborhood_of': ['location', 'location'],
            'person-place_of_birth': ['person', 'location'],
            },
    'chinese': {'所属行业': ['企业', '行业'], '行业上级': ['行业', '行业'], '上游原材料': ['产品', '原材料'], '下游产品': ['产品', '产品'], '公司主营': ['企业', '产品'], '公司主营': ['企业', '原材料'],
                '产品小类': ['产品', '产品']}
}

func.push({
    name: "根据内容出图谱",
    description: "根据内容出知识图谱，目前以产业链知识图谱为核心",
    question: async () => {
        app.chat = [{ "role": "user", "content": re_s1_p.chinese.replace('{}', "与传统悬架工作原理相一致的是，空气弹簧、 CDC 减振器之间也是相辅相成的关系，\n" +
                "空气弹簧能够通过簧内气压变化来改变车身的高度以及弹簧刚度，CDC 减振器可以通过阻\n" +
                "尼系数的连续变化来调节悬架软硬，两者共同作用则能够提升驾乘的舒适性和操控性。") },
            { "role": "AI", "content": '(上游原材料)' }]
        lsdh(true)//打开历史对话
        Q = app.question
        //app.max_length = 4096
        resp = (await send(re_s1_p.chinese.replace('{}', Q) , Q))
        let str1 = resp.replaceAll("(", "");
        let str2 = str1.replaceAll(")", "");
        let str3 = str2.replaceAll("，", ",");
        let str4 = str3.replaceAll("\"", "");
        console.log("str4=====",str4)
        res1 = str4.split(",")
        console.log("res1=====",res1)
        content = ["------------------------------正文------------------------------"]
        resp1 = []
        if (res1!=[]){
            for (let i in res1) {
                let line = res1[i].trim()
                console.log("line=====",line)
                let rels = df_ret.chinese[line]
                console.log("rels=====",rels)
                resp3 = ''
                app.chat = []
                app.chat.push({ "role": "user", "content": re_s3_p.chinese.replace('{}', "大众汽车生产的悬架与传统悬架工作原理相一致的是，空气弹簧、 CDC 减振器之间也是相辅相成的关系，\n" +
                    "空气弹簧能够通过簧内气压变化来改变车身的高度以及弹簧刚度，CDC 减振器可以通过阻\n" +
                    "尼系数的连续变化来调节悬架软硬，两者共同作用则能够提升驾乘的舒适性和操控性。").replace('{}', '产品，原材料，企业') },
                { "role": "AI", "content": '{"产品":["悬架"],"原材料":["空气弹簧","CDC 减振器"],"企业":["大众汽车"]}\n' })
                if (rels != undefined){
                    console.log("打印数据：",re_s3_p.chinese.replace('{}', Q).replace('{}', rels[0]+'，'+rels[1]))
                    resp3 = (await send(re_s3_p.chinese.replace('{}', Q).replace('{}', rels[0]+'，'+rels[1]) + "\n\n"))
                }
                app.chat = []
                app.chat.push({ "role": "user", "content": re_s2_p.chinese.replace('{}', "大众汽车生产的悬架与传统悬架工作原理相一致的是，空气弹簧、 CDC 减振器之间也是相辅相成的关系，\n" +
                    "空气弹簧能够通过簧内气压变化来改变车身的高度以及弹簧刚度，CDC 减振器可以通过阻\n" +
                    "尼系数的连续变化来调节悬架软硬，两者共同作用则能够提升驾乘的舒适性和操控性。").replace('{}', '产品').replace('{}', '原材料').replace('{}', '上游原材料').replace('{}', '(悬架,空气弹簧,CDC 减振器)').replace('{}', '企业').replace('{}', '原材料') },
                { "role": "AI", "content": '| 产品 | 原材料 |\n|  --- | --- |\n| 悬架 | 空气弹簧 |\n| 悬架 | CDC 减振器 |\n' })
                if (rels != undefined){
                    console.log("打印数据：",re_s2_p.chinese.replace('{}', Q).replace('{}', rels[0]).replace('{}', rels[1]).replace('{}', line).replace('{}', resp3).replace('{}', rels[0]).replace('{}', rels[1]))
                    content.push(await send(re_s2_p.chinese.replace('{}', Q).replace('{}', rels[0]).replace('{}', rels[1]).replace('{}', line).replace('{}', resp3).replace('{}', rels[0]).replace('{}', rels[1]), rels) + "\n\n")
                }
            }
        }
        else{
            console.log("出问题了")
        }
        lsdh(false)//打开历史对话
        content = content.join("\n\n")
        add_conversation("user",  Q )
        add_conversation("AI",  content )
        console.log(content)
        copy(content)

    },
})