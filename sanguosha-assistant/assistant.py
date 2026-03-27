#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
三国杀问答助手
Sanguosha Assistant - 智能问答系统
"""

import json
import os
import re
from typing import List, Dict, Optional


class SanguoshaAssistant:
    """三国杀问答助手主类"""
    
    def __init__(self, data_dir: str = None):
        """初始化助手，加载数据"""
        if data_dir is None:
            data_dir = os.path.join(os.path.dirname(__file__), 'data')
        
        self.heroes = self._load_json(os.path.join(data_dir, 'heroes.json'))['heroes']
        self.cards = self._load_json(os.path.join(data_dir, 'cards.json'))
        self.rules = self._load_json(os.path.join(data_dir, 'rules.json'))
    
    def _load_json(self, filepath: str) -> dict:
        """加载JSON数据文件"""
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            return {}
    
    def ask(self, question: str) -> str:
        """
        回答三国杀相关问题
        
        Args:
            question: 用户问题
            
        Returns:
            回答内容
        """
        question = question.strip()
        
        # 检测问题类型并路由到对应的回答方法
        if self._is_hero_question(question):
            return self._answer_hero_question(question)
        elif self._is_card_question(question):
            return self._answer_card_question(question)
        elif self._is_rule_question(question):
            return self._answer_rule_question(question)
        else:
            return self._general_answer(question)
    
    def _is_hero_question(self, question: str) -> bool:
        """判断是否为武将相关问题"""
        keywords = ['武将', '技能', '技能是什么', '什么技能', '势力', '体力', '主公', '忠臣', '反贼', '内奸']
        hero_names = [h['name'] for h in self.heroes]
        
        # 检查是否包含武将名或相关关键词
        for name in hero_names:
            if name in question:
                return True
        for kw in keywords:
            if kw in question:
                return True
        return False
    
    def _is_card_question(self, question: str) -> bool:
        """判断是否为卡牌相关问题"""
        keywords = ['牌', '卡牌', '锦囊', '装备', '武器', '防具', '马']
        card_names = ['杀', '闪', '桃', '决斗', '过河拆桥', '顺手牵羊', '无中生有', '南蛮入侵', '万箭齐发', '乐不思蜀', '闪电', '五谷丰登', '桃园结义', '诸葛连弩', '青釭剑', '丈八蛇矛', '贯石斧', '青龙偃月刀', '八卦阵', '仁王盾']
        
        # 先检查是否直接提到卡牌名
        for name in card_names:
            if name in question:
                return True
        
        # 再检查关键词
        for kw in keywords:
            if kw in question:
                return True
        return False
    
    def _is_rule_question(self, question: str) -> bool:
        """判断是否为规则相关问题"""
        keywords = ['规则', '怎么玩', '怎么使用', '判定', '距离', '濒死', '死亡', '阶段', '流程', '回合']
        for kw in keywords:
            if kw in question:
                return True
        return False
    
    def _answer_hero_question(self, question: str) -> str:
        """回答武将相关问题"""
        # 尝试匹配武将名
        for hero in self.heroes:
            if hero['name'] in question:
                return self._format_hero_info(hero)
        
        # 如果没有匹配到具体武将，返回武将列表
        hero_list = [h['name'] for h in self.heroes]
        return f"目前收录的武将有：{', '.join(hero_list)}\n请问你想了解哪位武将？"
    
    def _format_hero_info(self, hero: dict) -> str:
        """格式化武将信息"""
        skills_text = '\n'.join([
            f"  【{s['name']}】({s['type']})：{s['description']}"
            for s in hero['skills']
        ])
        
        return f"""【{hero['name']}】{hero['title']}
势力：{hero['faction']} | 体力：{hero['health']}

技能：
{skills_text}"""
    
    def _answer_card_question(self, question: str) -> str:
        """回答卡牌相关问题"""
        # 清理问题中的特殊字符
        clean_question = question.replace('【', '').replace('】', '').replace('[', '').replace(']', '')
        
        # 检查基本牌
        for card in self.cards.get('basic_cards', []):
            if card['name'] in clean_question or clean_question in card['name']:
                return self._format_card_info(card)
        
        # 检查锦囊牌
        for card in self.cards.get('trick_cards', []):
            if card['name'] in clean_question or clean_question in card['name']:
                return self._format_card_info(card)
        
        # 检查装备牌
        for card in self.cards.get('equipment_cards', []):
            if card['name'] in clean_question or clean_question in card['name']:
                return self._format_card_info(card)
        
        # 列出卡牌类型
        return """卡牌分为三大类：
1. 基本牌：杀、闪、桃
2. 锦囊牌：决斗、过河拆桥、顺手牵羊、无中生有、南蛮入侵、万箭齐发、乐不思蜀、闪电等
3. 装备牌：武器、防具、马匹

请问你想了解哪种卡牌？"""
    
    def _format_card_info(self, card: dict) -> str:
        """格式化卡牌信息"""
        result = f"""【{card['name']}】{card['type']}
效果：{card['description']}"""
        
        if 'notes' in card:
            result += f"\n备注：{card['notes']}"
        
        if 'attack_range' in card:
            result += f"\n攻击范围：{card['attack_range']}"
        
        return result
    
    def _answer_rule_question(self, question: str) -> str:
        """回答规则相关问题"""
        # 检查FAQ
        for faq in self.rules.get('faq', []):
            if any(kw in question for kw in faq['question']):
                return f"Q: {faq['question']}\nA: {faq['answer']}"
        
        # 检查基本规则
        for rule in self.rules.get('basic_rules', []):
            if rule['title'] in question:
                return f"【{rule['title']}】\n{rule['content']}"
        
        # 返回规则概述
        rules_text = '\n'.join([
            f"- {r['title']}"
            for r in self.rules.get('basic_rules', [])
        ])
        
        return f"""游戏规则概述：
{rules_text}

请问你想了解哪个具体规则？"""
    
    def _general_answer(self, question: str) -> str:
        """通用回答"""
        return """你好！我是三国杀问答助手，可以帮你解答以下问题：

🎴 武将查询：告诉我武将名，我告诉你技能效果
  例："诸葛亮的技能是什么？"

🃏 卡牌查询：告诉我卡牌名，我告诉你使用方法
  例："【南蛮入侵】怎么用？"

📜 规则解答：询问游戏规则
  例："判定是什么？"

请问你想了解什么？"""
    
    def search_hero(self, name: str) -> Optional[dict]:
        """搜索武将"""
        for hero in self.heroes:
            if name in hero['name'] or hero['name'] in name:
                return hero
        return None
    
    def search_card(self, name: str) -> Optional[dict]:
        """搜索卡牌"""
        all_cards = (
            self.cards.get('basic_cards', []) +
            self.cards.get('trick_cards', []) +
            self.cards.get('equipment_cards', [])
        )
        for card in all_cards:
            if name in card['name'] or card['name'] in name:
                return card
        return None


def interactive_mode():
    """交互模式"""
    print("=" * 50)
    print("   三国杀问答助手 - 交互模式")
    print("=" * 50)
    print("输入问题进行查询，输入 'quit' 或 'exit' 退出\n")
    
    assistant = SanguoshaAssistant()
    
    while True:
        try:
            question = input("👤 你: ").strip()
            
            if question.lower() in ['quit', 'exit', '退出', 'q']:
                print("\n再见！祝你游戏愉快！🎮")
                break
            
            if not question:
                continue
            
            answer = assistant.ask(question)
            print(f"\n🤖 助手: {answer}\n")
            
        except KeyboardInterrupt:
            print("\n\n再见！祝你游戏愉快！🎮")
            break


if __name__ == '__main__':
    interactive_mode()
