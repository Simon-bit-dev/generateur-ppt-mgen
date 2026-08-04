from __future__ import annotations

from typing import Any

from .models import PlaceholderValue as PH
from .models import SlideSpec


def footer_left(data: dict[str, Any]) -> str:
    support_format = data["presentation"]["format"]
    prefix = {
        "CONFÉRENCE": "Conférence",
        "WEBINAIRE": "Webinaire",
        "FORMAT À PRÉCISER": "Format à préciser",
    }[support_format]
    return f"{prefix} : {data['presentation']['title']}"


def build_conference_specs(data: dict[str, Any]) -> list[SlideSpec]:
    p = data["presentation"]
    opening = data["opening"]
    sections = data["sections"]
    left = footer_left(data)
    date = p["date"]

    specs: list[SlideSpec] = [
        SlideSpec("01_Couverture", "Couverture", (
            PH(23, p["format"]),
            PH(24, p["title"]),
            PH(25, p["subtitle"], required=False),
            PH(26, date),
        ), clone_slide_number=False),
        SlideSpec("02_Bienvenue", "Bienvenue", (PH(22, opening["welcome_sentence"]),), clone_slide_number=False),
        SlideSpec("03_Intervenants", "Intervenants", (PH(21, opening["speaker_identification"]), PH(25, left), PH(26, date))),
        SlideSpec("04_Sommaire", "Sommaire", (PH(25, left), PH(26, date))),
        SlideSpec("05_Transition_Accueil", "Transition Accueil", (PH(24, sections["accueil_objectifs"]["transition_sentence"]),), clone_slide_number=False),
        SlideSpec("06_Objectifs", "Objectifs", (PH(24, sections["accueil_objectifs"]["objectives"]), PH(25, left), PH(26, date))),
        SlideSpec("07_Brise_glace", "Brise-glace", (
            PH(24, sections["accueil_objectifs"]["icebreaker"]["question"]),
            PH(25, sections["accueil_objectifs"]["icebreaker"]["instruction"]),
            PH(28, left),
            PH(27, date),
        )),
        SlideSpec("08_Transition_Repères_théoriques", "Transition Repères théoriques", (PH(24, sections["reperes_theoriques"]["transition_sentence"]),), clone_slide_number=False),
    ]

    for item in sections["reperes_theoriques"]["slides"]:
        if item["enabled"]:
            specs.append(SlideSpec("09_Corps_Repères_théoriques", f"Repère {item['slot']}", (
                PH(25, item["title"]),
                PH(26, item["body"], "bullet_text"),
                PH(29, left),
                PH(28, date),
            )))

    specs.append(SlideSpec("10_Transition_Conseils_pratiques", "Transition Conseils pratiques", (PH(24, sections["conseils_pratiques"]["transition_sentence"]),), clone_slide_number=False))
    for item in sections["conseils_pratiques"]["slides"]:
        if item["enabled"]:
            specs.append(SlideSpec("11_Corps_Conseils_pratiques", f"Conseil {item['slot']}", (
                PH(25, item["title"]),
                PH(26, item["body"], "bullet_text"),
                PH(29, left),
                PH(28, date),
            )))

    specs.append(SlideSpec("12_Transition_Testez_connaissances", "Transition Testez vos connaissances", (), clone_slide_number=False))
    for question in sections["testez_connaissances"]["questions"]:
        specs.append(SlideSpec("13_Vrai_ou_Faux", f"Question {question['number']}", (
            PH(21, question["statement"]),
            PH(25, left),
            PH(26, date),
        )))
        if question["verdict"] == "VRAI":
            specs.append(SlideSpec("15_Vrai", f"Correction {question['number']} — Vrai", (
                PH(21, question["correction_message"]),
                PH(27, question["explanation"], "bullet_text"),
                PH(25, left),
                PH(26, date),
            )))
        else:
            specs.append(SlideSpec("14_Faux", f"Correction {question['number']} — Faux", (
                PH(25, question["correction_message"]),
                PH(26, question["explanation"], "bullet_text"),
                PH(29, left),
                PH(28, date),
            )))

    specs.append(SlideSpec("16_Transition_Passer_action", "Transition Passer à l’action", (PH(24, sections["passer_action"]["transition_sentence"]),), clone_slide_number=False))
    for item in sections["passer_action"]["slides"]:
        if item["enabled"]:
            specs.append(SlideSpec("17_Corps_Passer_action", f"Passer à l’action {item['slot']}", (
                PH(25, item["title"]),
                PH(26, item["body"], "numbered_text"),
                PH(29, left),
                PH(28, date),
            )))

    specs.extend([
        SlideSpec("18_Transition_Conclusion", "Transition Conclusion", (PH(24, sections["conclusion"]["transition_sentence"]),), clone_slide_number=False),
        SlideSpec("19_Synthèse", "Synthèse", (PH(26, sections["conclusion"]["synthesis"], "bullet_text"), PH(25, left), PH(27, date))),
        SlideSpec("20_Pour_aller_plus_loin", "Pour aller plus loin", (PH(26, sections["conclusion"]["further_resources"], "bullet_text"), PH(25, left), PH(27, date))),
        SlideSpec("21_Vivoptim", "Vivoptim", (PH(25, left), PH(26, date))),
        SlideSpec("22_4ème_couverture", "Quatrième de couverture", (PH(26, f"{left}\n{date}"),)),
    ])
    return specs
