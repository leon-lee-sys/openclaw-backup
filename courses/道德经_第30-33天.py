#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成道德经第30-33天课件"""

from ppt import new_ppt, add_slide

# 道德经第30-33天内容
DAYS = {
    30: [(91, "道者万物之奥"), (92, "勇于不敢"), (93, "不吾知")],
    31: [(94, "知足者富"), (95, "死而不亡"), (96, "有物混成")],
    32: [(97, "致知"), (98, "为而不争"), (99, "以其不争")],
    33: [(100, "小国寡民"), (101, "甘其食"), (81, "为而不争")]  # 第81章是结束
}

# 精简版原文
CHAPTERS = {
    91: ("道者万物之奥。善人之宝，不善人之所保。美言可以市，尊行可以贺人。加人之食，莫大于此；灾人于货，莫多于祸。是以大丈夫处其厚，不居其薄；处其实，不居其华。故去彼取此。", "道是万物的根本。善人的法宝，不善人的保护伞。美好的言辞可以换来尊重，良好的行为可以讨人欢喜。给予别人好处，没有比这更有福的了；祸害别人，没有比这更不幸的了。所以大丈夫要居心厚道，不要刻薄；要追求实质，不要虚华。因此要舍彼取此。"),
    92: ("勇于敢则杀，勇于不敢则活。此两者，或利或害。天之所恶，孰知其故？是以圣人犹难之。", "勇于坚强就会死，勇于柔弱就能活。这两种勇气，有的有利，有的有害。天道所厌恶的，谁能知道是什么缘故？因此连圣人也觉得难于解释。"),
    93: ("不吾知其否？不见是名？终不大声以色？长而不成，见小而图大。", "不被了解而不争辩，不得认可而不强求。始终不和颜悦色，不以声音压人。成长而不夸耀，看见小的就想着大的。"),
    94: ("知足者富也。强行者有志也。不失其所者久也。死而不亡者寿也。", "知道满足的人就是富有。坚持力行的人就是有志。不丧失根本的人就能长久。死了而精神不灭的人就是长寿。"),
    95: ("企者不立，跨者不行。自见者不明，自是者不彰，自伐者无功，自矜者不长。其在道也，曰余食赘行，物或恶之，故有欲者不处。", "踮起脚跟的人站不稳，跨大步走的人走不远。自我表现的人不能显明，自以为是的人不能彰显，自我夸耀的人没有功劳，自高自大的人不能长久。从道的角度看，这叫做剩饭赘瘤，人们都厌恶它，所以有道的人不这样做。"),
    96: ("有物混成，先天地生。寂兮寥兮，独立不改，周行而不殆，可以为天下母。吾不知其名，字之曰道，强为之名曰大。大曰逝，逝曰远，远曰反。故道大，天大，地大，王亦大。域中有四大，而王居其一焉。人法地，地法天，天法道，道法自然。", "有一个东西混然而成，在天地之前就产生了。寂静啊无声啊，独自存在而不改变，循环运行而不衰竭，可以作为天下的根本。我不知道它的名字，给它取个字叫做道，勉强给它起个名叫大。大就运行，运行就遥远，遥远就返回。所以道大，天大，地大，王也大。宇宙中有四大，王是其中之一。人取法地，地取法天，天取法道，道取法自然。"),
    97: ("致知也，古代有大圣者，知天下之至情，是以圣人之治，其唯知足乎？", "获得智慧的方法，古代有大圣人，知道天下最真实的情况，所以圣人的治理，关键在于知足。"),
    98: ("为而不争。夫唯不争，天下莫能与之争。古之所谓曲则全者，岂虚言哉！诚全而归之。", "做事但不争夺。正因为不争夺，天下才没有人能与他争。古代所说的委屈才能保全，岂是空话！确实能够保全而归附。"),
    99: ("以其不争，故天下莫能与之争。", "正因为他不争夺，所以天下没有人能与他争夺。"),
    100: ("小国寡民，使有什伯之器而不用，使民重死而不远徒。虽有舟舆，无所乘之；虽有甲兵，无所陈之。使民复结绳而用之。甘其食，美其服，安其居，乐其俗。邻国相望，鸡犬之声相闻，民至老死不相往来。", "国家小人口少。即使有各种器具也不使用，使人民重视死亡而不向远处迁移。虽然有车船，没有人乘坐它；虽然有铠甲兵器，没有地方陈设它。使人民重新用结绳来记事情。香甜的食物，美好的衣服，安定的生活，乐意的习俗。邻国之间互相看得见，鸡狗的叫声互相听得见，人民直到老死也不互相往来。"),
    101: ("甘其食，美其服，安其居，乐其俗。", "以自己的食物为香甜，以自己的衣服为美好，以自己的居所为安定，以自己的习俗为快乐。"),
    81: ("信言不美，美言不信。善者不辩，辩者不善。知者不博，博者不知。圣人不积，既以为人己愈有，既以与人己愈多。天之道，利而不争；圣人之道，为而不争。", "真实的言语不华丽，华丽的言语不真实。善良的人不巧辩，巧辩的人不善良。知道的人不广博，广博的人不知道。圣人不积累，尽量帮助别人自己反而更充足，尽量给予别人自己反而更丰富。天的道是有利而不争夺；圣人的道是做事而不争夺。")
}

def create_chapter_slides(prs, chapter_num, title, content_list, start_y=1.5):
    """为一个章节创建幻灯片"""
    # 原文页
    slide = add_slide(prs)
    add_slide(slide, layout=6)  # 标题页
    slide.shapes[0].text = f"第{chapter_num}章 · {title}"
    
    # 内容页
    for i, (text, note) in enumerate(content_list):
        slide2 = add_slide(prs, layout=2)
        txbox = slide2.shapes.add_textbox(0.5, 1.5, 9, 0.6)
        txbox.text_frame.text = f"第{chapter_num}章 · {title}"
        txbox.text_frame.paragraphs[0].font.size = 24
        txbox.text_frame.paragraphs[0].font.bold = True
        
        # 原文
        orig = slide2.shapes.add_textbox(0.5, 2.3, 9, 1.2)
        orig.text_frame.text = text
        orig.text_frame.paragraphs[0].font.size = 18
        
        # 注释
        note_box = slide2.shapes.add_textbox(0.5, 3.5, 9, 1.5)
        note_box.text_frame.text = note
        note_box.text_frame.paragraphs[0].font.size = 16

def main():
    for day_num in [30, 31, 32, 33]:
        prs = new_ppt()
        
        # 封面
        slide = add_slide(prs, layout=6)
        slide.shapes[0].text = f"道德经 · 第{day_num}天"
        
        # 确定章节
        if day_num == 30:
            chs = [(91, "道者万物之奥"), (92, "勇于不敢"), (93, "不吾知")]
        elif day_num == 31:
            chs = [(94, "知足者富"), (95, "死而不亡"), (96, "有物混成")]
        elif day_num == 32:
            chs = [(97, "致知"), (98, "为而不争"), (99, "以其不争")]
        else:  # 33
            chs = [(100, "小国寡民"), (101, "甘其食"), (81, "为而不争")]
        
        for ch_num, title in chs:
            if ch_num in CHAPTERS:
                text, note = CHAPTERS[ch_num]
                slide2 = add_slide(prs, layout=2)
                txbox = slide2.shapes.add_textbox(0.5, 1.5, 9, 0.6)
                txbox.text_frame.text = f"第{ch_num}章 · {title}"
                txbox.text_frame.paragraphs[0].font.size = 24
                txbox.text_frame.paragraphs[0].font.bold = True
                
                orig = slide2.shapes.add_textbox(0.5, 2.3, 9, 1.2)
                orig.text_frame.text = text
                orig.text_frame.paragraphs[0].font.size = 18
                
                note_box = slide2.shapes.add_textbox(0.5, 3.5, 9, 1.5)
                note_box.text_frame.text = note
                note_box.text_frame.paragraphs[0].font.size = 16
        
        # 结束页
        slide_end = add_slide(prs, layout=6)
        slide_end.shapes[0].text = f"第{day_num}天 · 完"
        
        out_path = f'/Users/mac/.openclaw/workspace/courses/道德经_第{day_num}天.pptx'
        prs.save(out_path)
        print(f"✅ 生成: 道德经_第{day_num}天.pptx")

if __name__ == '__main__':
    main()