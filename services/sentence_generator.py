# -*- coding: utf-8 -*-
"""
template phrase composing service
generate automaticly examples including the words
"""

import random
from typing import List, Dict

class TemplateSentenceGenerator:
    def __init__(self):
        # templates of different levels
        self.templates = {
            "beginner": [
                "I like {word}.",
                "This is a {word}.",
                "My {word} is red.",
                "The {word} is big.",
                "I have a {word}."
            ],
            "intermediate": [
                "I need to buy some {word} for dinner.",
                "The {word} book is very interesting.",
                "She works as a {word} in the company.",
                "We should discuss this {word} tomorrow.",
                "This {word} method is quite effective."
            ],
            "advanced": [
                "The implementation of this {word} algorithm requires careful consideration.",
                "Despite the {word} challenges, the project was completed successfully.",
                "The researcher's {word} analysis provided valuable insights.",
                "Companies are increasingly adopting {word} strategies to improve efficiency.",
                "The {word} phenomenon has significant implications for future development."
            ]
        }
        
        # special rules for verbs
        self.special_rules = {
            "be": ["I {word} happy.", "She {word} a teacher.", "They {be} students."],
            "have": ["I {word} a car.", "We {have} two cats.", "She {have} experience."],
            # TODO add more special unregular verbs
        }
    
    def generate_sentence(self, word: str, level: str = "beginner") -> str:
        """
        generate phrase including word input
        :param word: target word
        :param level: difficulty
        :return: generated phrase
        """
        # special grammar
        if word.lower() in self.special_rules:
            templates = self.special_rules[word.lower()]
            template = random.choice(templates)
            return template.format(**{word.lower(): word})
        
        # normal template
        if level not in self.templates:
            level = "beginner"
            
        templates = self.templates[level]
        template = random.choice(templates)
        return template.format(word=word)
    
    def generate_multiple_sentences(self, word: str, count: int = 3, 
                                  levels: List[str] = None) -> List[str]:
        """
        random difficulty
        :param word:
        :param count: total number
        :param levels:
        :return: example list
        """
        if levels is None:
            levels = ["beginner", "intermediate", "advanced"]
        
        sentences = []
        used_templates = set()
        
        for i in range(count):
            level = levels[i % len(levels)] if levels else "beginner"
            
            # ensure to avoid duplicates
            available_templates = [
                t for t in self.templates.get(level, []) 
                if t not in used_templates
            ]
            
            if not available_templates:
                # if all templates have been used, reset
                used_templates.clear()
                available_templates = self.templates.get(level, [])
            
            if available_templates:
                template = random.choice(available_templates)
                used_templates.add(template)
                sentence = template.format(word=word)
                sentences.append(sentence)
        
        return sentences
