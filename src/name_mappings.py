NAME_MAP = {
    "Assuncao Gambetta Clemente, Fernanda": [
        "Assuncao Gambetta Clemente, Fernanda",
        "Fernanda Assuncao Gambetta Clemente",
    ],
    "Magro, Daniel": [
        "Magro, Daniel",
    ],
    "Pires Rosa, Claudia": [
        "Pires Rosa, Claudia",
        "Claudia Pires Rosa",
        "Rosa, Claudia (ext)",
        "Pires Rosa, Claudia (ext)",
        "Claudia Rosa",
    ],
    "Helbing, Björn": [
        "Helbing, Björn",
        "Helbing, Bjoern",
        "B. Helbing",
        "Björn Helbing",
    ],
    "Matos Oliveira, Ana Rita": [
        "Matos Oliveira, Ana Rita",
        "Ana Rita Matos Oliveira",
        "Matos dos Santos Oliveira, Ana Rita",
        "Matos Oliveira, Rita",
    ],
    "Pires, Filipe": [
        "Pires, Filipe",
        "Guerreiro Luis Pires, Filipe Viegas",
        "Filipe Pires",
    ],
    "Plácido, Andreia": [
        "Plácido, Andreia",
        "Moreira Cristo Placido, Andreia Sofia",
        "Andreia Plácido",
        "Andreia Sofia Moreira Cristo Placido",
    ],
    "Antunes, Ricardo": [
        "Antunes, Ricardo",
        "Ricardo Antunes",
    ],
    "Fernandes Redondo, Amanda": [
        "Fernandes Redondo, Amanda",
        "Amanda Fernandes Redondo",
        "Fernandez Redondo, Amanda",
    ],
    "Zouine, Meryem": [
        "Zouine, Meryem",
        "Meryem Zouine",
    ],
    "Cerezo, Alberto": [
        "Cerezo, Alberto",
        "Cerezo Ruiz, Alberto",
        "Alberto Cerezo",
    ],
    "Swoboda, Claudia": [
        "Swoboda, Claudia",
        "Claudia Swoboda",
    ],
    "Bicho, Rita": [
        "Bicho, Rita",
        "Lazaro Bicho, Rita Sofia",
        "Rita Bicho",
        "Rita Sofia Lazaro Bicho",
    ],
    "Lopes Fonseca, Mario Andre": [
        "Lopes Fonseca, Mario Andre",
        "Mario Andre Lopes Fonseca",
    ],
    "Wiesheu, Andreas": [
        "Wiesheu, Andreas",
        "Andreas Wiesheu",
    ],
    "Heldwein, Christian": [
        "Heldwein, Christian",
        "Christian Heldwein",
    ],
    "Hernandes Vaz, Joao Rafael": [
        "Hernandes Vaz, Joao Rafael",
        "Joao Rafael Hernandes Vaz",
    ],
    "Vitorino, Diana": [
        "Vitorino, Diana",
        "Diana Vitorino",
    ],
    "Candeias Gracioso, Sara Margarida": [
        "Candeias Gracioso, Sara Margarida",
        "Sara Margarida Candeias Gracioso",
    ],
    "do Nascimento Matos Manso, Rui Pedro": [
        "do Nascimento Matos Manso, Rui Pedro",
        "Rui Pedro do Nascimento Matos Manso",
    ],
}


ALIAS_TO_CANONICAL = {}
for canonical, aliases in NAME_MAP.items():
    for alias in aliases:
        ALIAS_TO_CANONICAL[alias.strip().lower()] = canonical

def normalize_name(name):
    if not name:
        return None
    return ALIAS_TO_CANONICAL.get(name.strip().lower(), name.strip())

# rows to skip 
SKIP_NAMES = [
        "Test - ARE 5240", "Service Management - ARE 5290 + int.", "JCC",
        "CHCM (Eviden)", "AWS Encryption Key (KMS for Eightfold; GBS)",
        "DirX (Ext)/(SAG Global)", "Integrations DPS", "GBS total",
        "TRE (Tupu) PO", "Eightfold (PO Q1/2025: 9708791569; Q2-Q4: tbd.)",
        "Provider", "CERT check (pen GBS)", "Accessibility Test",
        "Travel & Hospitality", "Additional costs", "Total Costs",
        "Accumulated costs Eightfold Crew", "Service Management - ARE 5290 + int.",
        "DirX PT -  5240", "DirX (Eviden)", "TRE (Tupu) (PO tbd)", "Eightfold",
        "Avature DT (PO 9708872111)", "Avature AM (PO 9708872111)",
        "Avature Healthcheck (PO", "Total costs", "Accumulated costs Avature Crew",
        "CHCM (Evdien)", "Avature ext. Careers Portal (PO 9709043748) - PDP",
        "Avature ext. Careers Portal (PO 9709132497)", "Travel & Hospitality ext. Provider",
        "Travel & Hospitality DE", "Travel & Hospitality ES / CZ / PT",
        "Total costs", "Accumulated costs Ext. Careers Port Crew", "Accumulated costs Preboarding",
        "Service Management","Total","Avature","DPS Internal Umlage (global)"
    ]