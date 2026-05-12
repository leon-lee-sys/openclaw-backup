#!/usr/bin/env python3
"""Generate Advanced Aesop's Fables English Learning Document L31-35"""

import sys
sys.path.insert(0, '/Users/mac/.openclaw/workspace/skills/docx-generator')

from scripts.docx_generator import DocxGenerator

# Define 5 advanced Aesop's Fables with full content
fables = [
    {
        "title": "The Town Mouse and the Country Mouse",
        "title_cn": "城里老鼠和乡下老鼠",
        "story_en": """A Town Mouse, who had relatives in the country, invited his cousin to come and see him. "Come," said he, "you shall eat the finest food and drink the richest wines, and you will escape the hardships of your simple life in the fields."

The Country Mouse, who had never tasted such abundance, eagerly accepted the invitation. When they reached the city, the Town Mouse led his guest through many winding streets to a grand house. There they found a table spread with the remains of a feast—delicate pastries, cold meats, honey cakes, and fruit of every kind.

But just as they were about to eat, the door of the room opened, and the servants entered to clear away the dishes. The two mice had barely time to escape to a hiding-place before the master of the house arrived with several guests.

"My dear friend," said the Country Mouse, as they crouched trembling behind the wainscoting, "I see now why you prefer a humble life in the country. This abundance of yours is accompanied by constant fear and danger. Give me my barley field and my humble burrow, where I can eat in peace and sleep without dread."

He who risks everything for everything may lose everything.""",
        "story_cn": """一只城里的老鼠有几个乡下的亲戚，他邀请表兄来看望他。"来吧，"他说，"你可以吃到最精美的食物，喝到最醇厚的美酒，还可以免去乡下简朴生活中的艰辛。"

从未尝过如此丰盛食物的乡下老鼠，欣然接受了邀请。当他们到达城市后，城里老鼠带着客人穿过许多弯曲的街道，来到一座宏伟的宅邸。在那里，他们发现桌上摆着宴席的残羹——精致的糕点、冷盘肉、蜂蜜蛋糕和各种水果。

可是正当他们要吃的时候，房间的门开了，仆人们进来收拾碗碟。两只老鼠刚刚逃到一个藏身之处，房子的主人就带着几位客人进来了。

"我亲爱的朋友，"当他们蜷缩在护壁板后面瑟瑟发抖时，乡下老鼠说道，"我现在明白你为什么更喜欢乡下简朴的生活了。你的这些丰盛食物伴随着不断的恐惧和危险。还是给我那片大麦田和简陋的洞穴吧，在那里我可以安宁地吃东西，不必提心吊胆地睡觉。"

为了得到一切而冒一切风险的人，可能会失去一切。""",
        "vocabulary": [
            ("relatives", "/ˈrelətɪvz/", "n. 亲戚，家属"),
            ("humble", "/ˈhʌmbl/", "adj. 谦逊的；简朴的"),
            ("abundance", "/əˈbʌndəns/", "n. 充裕，丰富"),
            ("eagerly", "/ˈiːɡəli/", "adv. 热切地，渴望地"),
            ("grand", "/ɡrænd/", "adj. 宏伟的，壮丽的"),
            ("delicate", "/ˈdelɪkət/", "adj. 精致的，精美的"),
            ("winding", "/ˈwaɪndɪŋ/", "adj. 弯曲的，蜿蜒的"),
            ("wainscoting", "/ˈweɪnskɒtɪŋ/", "n. 护壁板，嵌板"),
            ("crouched", "/kraʊtʃt/", "v. 蜷缩，蹲伏"),
            ("trembling", "/ˈtremblɪŋ/", "adj. 颤抖的，发抖的"),
            ("dread", "/dred/", "n. 恐惧，忧虑"),
            ("barley", "/ˈbɑːli/", "n. 大麦"),
            ("burrow", "/ˈbʌrəʊ/", "n. 地洞，洞穴"),
            ("constant", "/ˈkɒnstənt/", "adj. 持续不断的，始终如一的"),
            ("accompany", "/əˈkʌmpəni/", "v. 陪伴，伴随"),
        ],
        "grammar": [
            "虚拟语气 (Subjunctive Mood): 'He who risks everything for everything may lose everything.' 这句话使用了含蓄条件句，表达了一种普遍真理。",
            "过去分词作状语: 'The Country Mouse, who had never tasted such abundance, eagerly accepted the invitation.' 描述先发生的动作。",
            "定语从句: 多个 'who' 和 'that' 引导的定语从句用于描述名词。",
            "宾语从句: 'I see why you prefer a humble life in the country.' 由 why 引导的宾语从句。",
        ],
        "moral": "真正的幸福不在于物质的丰富，而在于内心的安宁。追求奢华往往伴随着恐惧和不安，而简朴的生活反而能带来真正的平和与满足。",
        "oral_en": [
            "What would you choose: a luxurious but dangerous life, or a simple but peaceful one? Why?",
            "Do you think the Town Mouse was happy with his life in the city? Why or why not?",
            "Can you think of a situation where having more actually means having less happiness?",
            "How does this fable relate to modern life in big cities?",
            "What is the difference between being wealthy and being content?",
        ],
        "oral_cn": [
            "你会选择什么：奢华但危险的生活，还是简朴但安宁的生活？为什么？",
            "你认为城里老鼠对他在城市里的生活满意吗？为什么？",
            "你能想到一个拥有更多反而意味着更少幸福的情况吗？",
            "这个寓言如何与现代社会的大城市生活联系？",
            "富裕和满足有什么区别？",
        ],
    },
    {
        "title": "The Lion, Jupiter, and the Elephant",
        "title_cn": "狮子、朱庇特与大象",
        "story_en": """A Lion, growing old and weak, lay sighing in his den. He was too feeble to hunt for food, and he began to think of his past glory, when he had been called the King of Beasts.

"Alas!" he cried, "how hard it is to grow old! I remember when my roar used to terrify all the animals in the forest. Now I can barely find strength to warn away the jackals that steal my food."

So the Lion made his way to Jupiter's temple, and there, in the presence of the gods, he made a humble prayer: "O mighty Jupiter, father of gods and men, grant that the beasts may see me as I was, and fear me once more!"

Jupiter, looking down from his golden throne, answered: "Your prayer is heard, but you must accept the mark of your wish." And he caused great horns to grow from the Lion's head.

All the beasts, seeing this strange new creature, fled in terror—until the Elephant, who had not seen the Lion before, stopped to examine him.

"What are those peculiar growths on your head?" asked the Elephant, pointing with his trunk to the horns.

"They are my crowns of glory," replied the Lion, proudly.

"I think not," said the Elephant, shaking his head. "They look more like big veins. I have veins like those in my head, and no one calls them crowns of glory."

Pride makes us see ourselves as we wish to be, not as we truly are.""",
        "story_cn": """一头狮子日渐衰老虚弱，躺在洞穴里叹息。他太虚弱了，无法捕猎食物，他开始怀念过去的荣耀——那时他被称为百兽之王。

"唉！"他喊道，"变老真是艰难！我记得从前我的吼声曾让森林里所有的动物都胆战心惊。现在我几乎无法有力气赶走偷我食物的豺狼。"

于是狮子前往朱庇特的神殿，在众神面前，他谦卑地祈祷："哦，伟大的朱庇特，众神与人类之父，请让野兽们能够看到我曾经的样子，再次敬畏我吧！"

朱庇特从他的金色宝座上俯视，回答道："你的祈祷已被听到，但你必须接受你愿望的标记。"于是他让巨大的角从狮子的头上长出。

所有的野兽看到这个奇怪的新生物，都恐惧地逃跑了——直到从未见过狮子的 Elephant（大象）停下来仔细观察他。

"你头上那些奇怪的突起物是什么？"大象用鼻子指着角问道。

"这是我荣耀的王冠，"狮子骄傲地回答。

"我不这么认为，"大象摇着头说，"它们看起来更像是大血管。我头上也有这样的血管，但没有人称它们为荣耀的王冠。"

骄傲使我们看到自己希望成为的样子，而不是我们真正的样子。""",
        "vocabulary": [
            ("feeble", "/ˈfiːbl/", "adj. 虚弱的，微弱的"),
            ("den", "/den/", "n. 兽穴，窝点"),
            ("glory", "/ˈɡlɔːri/", "n. 荣耀，荣誉"),
            ("terrify", "/ˈterɪfaɪ/", "v. 使恐惧，使惊吓"),
            ("jackal", "/ˈdʒækl/", "n. 豺狼，豺"),
            ("mighty", "/ˈmaɪti/", "adj. 强大的，有力的"),
            ("throne", "/θrəʊn/", "n. 王座，御座"),
            ("peculiar", "/pɪˈkjuːliə/", "adj. 奇怪的，独特的"),
            ("vein", "/veɪn/", "n. 静脉，血管"),
            ("crown", "/kraʊn/", "n. 王冠，桂冠"),
            ("humble", "/ˈhʌmbl/", "adj. 谦逊的，卑微的"),
            ("prayer", "/preə/", "n. 祈祷，祷告"),
            ("flee (fled)", "/fliː/", "v. 逃跑，逃离"),
            ("examine", "/ɪɡˈzæmɪn/", "v. 检查，仔细观察"),
            ("trunk", "/trʌŋk/", "n. 象鼻，鼻子"),
        ],
        "grammar": [
            "感叹句: 'Alas! how hard it is to grow old!' 使用 how + adj. 的感叹句结构。",
            "宾语补足语: 'Jupiter caused great horns to grow from the Lion's head.' make/cause + 宾语 + 不定式结构。",
            "独立主格结构: 'the Elephant, who had not seen the Lion before, stopped...' 描述同时发生的动作。",
            "直接引语和间接引语: 多个直接引语用于对话中。",
            "谚语理解: 'Pride makes us see ourselves as we wish to be' — 抽象主语 + 动词第三人称单数形式。",
        ],
        "moral": "骄傲使我们自欺欺人，看不到自己的真实面目。真正的尊严不在于外表的变化或他人的敬畏，而在于接受自我、保持本色。虚荣的追求往往只会暴露我们的弱点。",
        "oral_en": [
            "Why do you think the Lion wanted the other animals to fear him again?",
            "What do you think the Elephant's response tells us about wisdom?",
            "Is there a difference between being respected and being feared? Explain.",
            "Have you ever wanted something that seemed good but turned out to be foolish?",
            "What does this fable teach us about accepting ourselves as we are?",
        ],
        "oral_cn": [
            "你认为狮子为什么想让其他动物再次敬畏他？",
            "你认为大象的回答告诉了我们什么关于智慧的道理？",
            "受人尊敬和令人恐惧有区别吗？请解释。",
            "你有没有曾经想要过某个看起来很好但实际上却很愚蠢的东西？",
            "这个寓言教给我们什么关于接受自我真实面目的道理？",
        ],
    },
    {
        "title": "The Eagle and the Arrow",
        "title_cn": "鹰与箭",
        "story_en": """A Hunter, who was practicing in the forest, let fly an Arrow at a Eagle that was soaring overhead. The Arrow found its mark, and the great bird fell to the earth with a broken wing.

As the Hunter hastened to secure his prize, he was astonished to see that the shaft of the Arrow was feathered with one of the Eagle's own plumes.

"It is a true saying," observed a spectator, "that we often give our feathers to our enemies to make arrows to wound us."

The Eagle, still living though wounded, looked at the arrow and recognized its own feather. "Ah," he said, with a bitter smile, "this is indeed a just punishment. I used to plume myself on being the swiftest and most majestic of birds. Now I am brought low by my own feathers."

Many men, in their pride, use their own cleverness to destroy themselves. The very gifts that make us great can become the instruments of our downfall if we are not careful.""",
        "story_cn": """一个猎人正在森林里练习射箭，他向一只在天上翱翔的老鹰射出一支箭。箭命中了目标，这只大鸟带着折断的翅膀落到地上。

当猎人急忙去获取他的猎物时，他惊讶地发现那支箭的箭杆上插着一根羽毛，正是老鹰自己的羽毛。

"有句俗话说得好，"一个旁观者说道，"我们常常把自己的羽毛送给敌人，让他们制成箭来伤害我们。"

老鹰虽然受了伤但还活着，他看着那支箭，认出了自己的羽毛。"啊，"他苦笑着说，"这确实是公正的惩罚。我过去常常以自己是所有鸟类中最快速、最雄伟的而沾沾自喜。现在我却栽在自己的羽毛上了。"

许多人在骄傲中用自己的聪明才智毁灭自己。如果我们不小心谨慎，那些使我们伟大的天赋可能成为我们衰败的工具。""",
        "vocabulary": [
            ("Arrow", "/ˈærəʊ/", "n. 箭"),
            ("Hunter", "/ˈhʌntə/", "n. 猎人"),
            ("soar", "/sɔː/", "v. 高飞，翱翔"),
            ("hasten", "/ˈheɪsn/", "v. 急忙，赶快"),
            ("astonished", "/əˈstɒnɪʃt/", "adj. 惊讶的，吃惊的"),
            ("shaft", "/ʃɑːft/", "n. 箭杆，轴"),
            ("plume", "/pluːm/", "n. 羽毛；v. 沾沾自喜"),
            ("spectator", "/spekˈteɪtə/", "n. 观众，旁观者"),
            ("majestic", "/məˈdʒestɪk/", "adj. 雄伟的，庄严的"),
            ("swiftest", "/ˈswɪftɪst/", "adj. 最快速的 (swift 最高级)"),
            ("cleverness", "/ˈklevənəs/", "n. 聪明，智慧"),
            ("downfall", "/ˈdaʊnfɔːl/", "n. 衰落，垮台"),
            ("just", "/dʒʌst/", "adj. 公正的，正义的"),
            ("punishment", "/ˈpʌnɪʃmənt/", "n. 惩罚，处罚"),
            ("instrument", "/ˈɪnstrəmənt/", "n. 工具，手段"),
        ],
        "grammar": [
            "现在分词作状语: 'The Arrow finding its mark' 描述结果。",
            "过去分词作定语: 'the shaft of the Arrow was feathered' — 被动完成式描述状态。",
            "谚语引用: 'It is a true saying...' 引入格言。",
            "反身代词: 'I used to plume myself on...' — 反身动词用法。",
            "让步状语从句: 'The very gifts that make us great can become...' 含蓄让步。",
        ],
        "moral": "我们最引以为傲的才能，有时会成为毁灭我们的工具。骄傲和自大会蒙蔽我们的判断，使我们成为自己成功的牺牲品。保持谦逊才能避免自食其果。",
        "oral_en": [
            "What do you think the spectator meant by giving 'our feathers to our enemies'?",
            "Can you think of examples in real life where someone's strength became their weakness?",
            "How can we avoid the Eagle's fate in using our own talents?",
            "What is the difference between being proud of something and being proud of yourself?",
            "Do you think the Eagle deserved his punishment? Why or why not?",
        ],
        "oral_cn": [
            "你认为旁观者说的把我们自己的羽毛送给敌人是什么意思？",
            "你能想到现实生活中一个人的优势变成劣势的例子吗？",
            "我们如何在运用自己的才能时避免重蹈鹰的覆辙？",
            "为某事感到骄傲和为自己感到骄傲有什么区别？",
            "你认为老鹰受到惩罚是应该的吗？为什么？",
        ],
    },
    {
        "title": "The Fowler and the Viper",
        "title_cn": "捕鸟人与毒蛇",
        "story_en": """A Fowler, as he traveled through a lonely part of the country, came upon a Viper lying in a sandy road, apparently frozen with cold. The kind-hearted man took up the reptile and placed it in his breast pocket to warm it with his own body heat.

When the Viper had recovered from the cold, it revived and, feeling itself at liberty, began to show its true nature. It bit its benefactor, and the poison quickly spread through his veins.

The Fowler, feeling the venom taking effect, said: "Alas! I have only myself to blame. Why did I take pity on such a wretch? The Viper's nature is to harm, not to be grateful. I should have remembered that kindness given to the wicked is always thrown away."

It is a foolish kindness that helps those who cannot appreciate it. A serpent cannot be tamed, and a thankless creature will always repay good with evil.""",
        "story_cn": """一个捕鸟人在穿越一片荒凉的乡村时，遇到一条毒蛇躺在沙路上，显然是被寒冷冻僵了。这个善良的人捡起这只爬行动物，把它放进胸前的口袋里，用自己的体温来温暖它。

当毒蛇从寒冷中恢复过来后，它苏醒过来，感到自己获得了自由，便开始显露出它的本性。它咬了它的恩人，毒液迅速蔓延到他的血管中。

捕鸟人感到毒液在起作用，说道："唉！我只能怪自己。我为什么要可怜这样的恶棍？毒蛇的本性就是伤害，而不是感恩。我应该记住，对恶人施以仁慈总是白费力气。"

帮助那些不懂得感激的人是一种愚蠢的善良。毒蛇无法驯服，不知感恩的生物总是以怨报德。""",
        "vocabulary": [
            ("Fowler", "/ˈfaʊlə/", "n. 捕鸟人，捕鸟者"),
            ("Viper", "/ˈvaɪpə/", "n. 毒蛇，蝰蛇"),
            ("sandy", "/ˈsændi/", "adj. 沙的，多沙的"),
            ("reptile", "/ˈreptaɪl/", "n. 爬行动物"),
            ("breast", "/brest/", "n. 胸部，乳房"),
            ("revive", "/rɪˈvaɪv/", "v. 苏醒，复活"),
            ("benefactor", "/ˈbenɪfæktə/", "n. 恩人，施惠者"),
            ("venom", "/ˈvenəm/", "n. 毒液，恶意"),
            ("wretch", "/retʃ/", "n. 可怜的人，恶棍"),
            ("serpent", "/ˈsɜːpənt/", "n. 大蛇，毒蛇"),
            ("wicked", "/ˈwɪkɪd/", "adj. 邪恶的，坏的"),
            ("thankless", "/ˈθæŋkləs/", "adj. 不知感恩的"),
            ("tamed", "/teɪmd/", "v. 驯服 (tame 过去式)"),
            ("repay", "/rɪˈpeɪ/", "v. 报答，偿还"),
            ("grateful", "/ˈɡreɪtfl/", "adj. 感激的，表示感谢的"),
        ],
        "grammar": [
            "现在分词作伴随状语: 'took up the reptile and placed it in his breast pocket' — 描述连续动作。",
            "过去分词作定语: 'a sandy road, apparently frozen with cold' — 描述状态。",
            "让步状语从句: 'apparently frozen with cold' — 独立主格结构作状语。",
            "直接引语: 捕鸟人临终的话语。",
            "名词性从句: 'Why did I take pity on such a wretch?' 问句作宾语。",
        ],
        "moral": "善良需要智慧相伴。对那些本性邪恶、不知感恩的人施以仁慈是愚蠢的。帮助他人之前，我们应该认清对方的本质，否则善良会被辜负，甚至招致祸患。",
        "oral_en": [
            "Do you think the Fowler was wrong to help the Viper? Why?",
            "What would you do if you found a dangerous animal in need of help?",
            "Is it possible to change someone's evil nature through kindness? Why or why not?",
            "How can we tell when kindness might be taken for granted?",
            "What is the difference between being kind and being wise?",
        ],
        "oral_cn": [
            "你认为捕鸟人帮助毒蛇错了吗？为什么？",
            "如果你发现一只危险的动物需要帮助，你会怎么做？",
            "通过善良有可能改变一个人的邪恶本性吗？为什么？",
            "我们如何判断善良何时可能被视为理所当然？",
            "善良和智慧有什么区别？",
        ],
    },
    {
        "title": "The Fisherman and the Little Fish",
        "title_cn": "渔夫与小鱼",
        "story_en": """A Fisherman, who lived near the sea, used to go out every day to catch fish. One morning, after casting his net several times without success, he drew it in and found only one small Fish swimming in it.

"Pray let me go, good Fisherman," said the Little Fish, struggling in the net. "I am far too small to be worth your cooking-pot. If you throw me back into the sea, I shall have time to grow large enough for your table some other day."

"Nay, little Fish," replied the Fisherman, "I never think of tomorrow when I can secure today's catch. A small fish that will be large tomorrow is not so valuable as a fish that is already in my hand."

It is unwise to let tomorrow's prospects make us neglect today's opportunities. A certainty in the hand is worth more than a hope in the bush. The fisherman who waits for a better catch often loses what he has already caught.""",
        "story_cn": """一个住在海边的渔夫，每天都出海捕鱼。一天早晨，他撒了几次网都没有收获，于是收网时发现网里只有一条小鱼在游来游去。

"求求你放了我吧，好心的渔夫，"小鱼在网里挣扎着说，"我实在太小了，不值得您把我放进锅里。如果您把我扔回海里，我就有时间长大，总有一天会长到足够您吃的。"

"不行，小鱼，"渔夫回答说，"当我现在就能抓住今天的收获时，我从不想明天的事。一条明天才会大的小鱼，不如已经在手里的一条鱼有价值。"

让明天的前景使我们忽视今天的机会是不明智的。手中确定的收获比虚无缥缈的希望更有价值。等待更好收获的渔夫往往会失去他已经抓住的东西。""",
        "vocabulary": [
            ("Fisherman", "/ˈfɪʃəmən/", "n. 渔夫，渔民"),
            ("cast (cast)", "/kɑːst/", "v. 投掷，撒（网）"),
            ("net", "/net/", "n. 网，网状物"),
            ("draw in", "/drɔː ɪn/", "phrasal v. 收网，拉近"),
            ("pray", "/preɪ/", "v. 请，求"),
            ("struggling", "/ˈstrʌɡlɪŋ/", "v. 挣扎，努力"),
            ("valuable", "/ˈvæljuəbl/", "adj. 有价值的，值钱的"),
            ("prospects", "/ˈprɒspekts/", "n. 前途，前景"),
            ("neglect", "/nɪˈɡlekt/", "v. 忽视，疏忽"),
            ("certainty", "/ˈsɜːtnti/", "n. 确定的事情，必然的事"),
            ("opportunity", "/ˌɒpəˈtjuːnəti/", "n. 机会，时机"),
            ("secure", "/sɪˈkjʊə/", "v. 获得，确保"),
            ("nay", "/neɪ/", "adv. 不，否"),
            ("cook-pot", "/kʊk pɒt/", "n. 锅，炖锅"),
            ("bush", "/bʊʃ/", "n. 灌木，灌木丛"),
        ],
        "grammar": [
            "条件状语从句: 'If you throw me back into the sea, I shall have time to grow...' — 第一条件句。",
            "比较级结构: 'A small fish that will be large tomorrow is not so valuable as a fish that is already in my hand.' — not so...as 比较结构。",
            "宾语从句: 'I never think of tomorrow when I can secure today's catch.' — 时间状语从句修饰 think。",
            "虚拟语气: 'I never think of tomorrow when I can secure today's catch.' — 事实陈述，与虚拟语气形成对比。",
            "谚语: 'A certainty in the hand is worth more than a hope in the bush.' — 简练的格言形式。",
        ],
        "moral": "不要为了追求更好的未来而放弃眼前确定的收获。机会来临时要紧紧抓住，不要总是寄希望于明天。今天的确定比明天的幻想更有价值。",
        "oral_en": [
            "Was the Fisherman wise to keep the small fish? Why or why not?",
            "Can you think of times when waiting for something better actually made you lose what you had?",
            "What's the difference between being patient and being too greedy?",
            "How does this fable apply to decision-making in business or career?",
            "Is it better to have a small certain gain or a chance at a larger one? Why?",
        ],
        "oral_cn": [
            "渔夫保留这条小鱼是明智的吗？为什么？",
            "你能想到有些时候等待更好的东西反而让你失去了已有的东西吗？",
            "有耐心和太贪婪有什么区别？",
            "这个寓言如何应用于商业或职业决策？",
            "获得确定的少量收益和有机会获得更大的收益，哪个更好？为什么？",
        ],
    },
]

# Create the document
gen = DocxGenerator()

# Set document header
gen.set_header_text("英语学习 L31-35 | 伊索寓言（高级）")

# Add main title
gen.add_title("英语学习 L31-35", level=1)
gen.add_title("伊索寓言（高级）Aesop's Fables — Advanced Level", level=2)
gen.add_spacing(1)

# Introduction
gen.add_paragraph("本课程包含5个高级伊索寓言故事，每个故事配有详细的双语原文、词汇表、语法解析、寓意分析和口语练习。建议逐课学习，反复朗读原文并完成口语练习。", font_size=11, italic=True)
gen.add_spacing(2)

# Generate each fable
for i, fable in enumerate(fables, start=31):
    # Lesson number
    gen.add_title(f"Lesson {i}: {fable['title']}", level=1)
    gen.add_title(fable['title_cn'], level=2)
    gen.add_spacing(1)

    # Story section
    gen.add_title("原文 | Original Text", level=3)
    gen.add_paragraph("【英文原文】", bold=True)
    gen.add_paragraph(fable['story_en'])
    gen.add_spacing(1)
    gen.add_paragraph("【中文译文】", bold=True)
    gen.add_paragraph(fable['story_cn'])
    gen.add_spacing(2)

    # Vocabulary section
    gen.add_title("词汇表 | Vocabulary", level=3)
    vocab_data = [["单词 Word", "音标 Pronunciation", "释义 Meaning"]]
    for word, pron, meaning in fable['vocabulary']:
        vocab_data.append([word, pron, meaning])
    gen.add_table(vocab_data)
    gen.add_spacing(2)

    # Grammar section
    gen.add_title("语法点解析 | Grammar Points", level=3)
    for point in fable['grammar']:
        gen.add_paragraph(f"• {point}")
    gen.add_spacing(2)

    # Moral section
    gen.add_title("寓意分析 | Moral Analysis", level=3)
    gen.add_paragraph(fable['moral'])
    gen.add_spacing(2)

    # Oral practice section
    gen.add_title("口语练习 | Oral Practice", level=3)
    gen.add_paragraph("【英文讨论问题】", bold=True)
    for j, q in enumerate(fable['oral_en'], start=1):
        gen.add_paragraph(f"{j}. {q}")
    gen.add_spacing(1)
    gen.add_paragraph("【中文讨论问题】", bold=True)
    for j, q in enumerate(fable['oral_cn'], start=1):
        gen.add_paragraph(f"{j}. {q}")
    gen.add_spacing(2)

    # Page break between lessons
    if i < 35:
        gen.add_page_break()

# Save the document
output_path = "/Users/mac/.openclaw/workspace/courses/英语学习_L31-35_AesopFables.docx"
gen.save(output_path)
print(f"Document saved to: {output_path}")
