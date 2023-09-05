import json
from dataclasses import dataclass, field
from enum import Enum, auto
from pathlib import Path
from typing import Optional

from PIL import Image, ImageChops

sources = Path("screenshots")
prefixes = sources.glob("*.png")
prefixes = {p.stem.split("_")[0] for p in prefixes}
prefixes = sorted(prefixes, key=int)

ranks = Path("ranks")
ranks.mkdir(exist_ok=True, parents=True)
ranks = ranks.glob("*.png")
ranks = [(int(p.stem.split("_")[0]), Image.open(p)) for p in ranks]

underlords = Path("underlords")
underlords.mkdir(exist_ok=True, parents=True)
underlords = underlords.glob("*.png")
underlords = [(p.stem.split("_")[0], Image.open(p)) for p in underlords]

heroes = Path("heroes")
heroes.mkdir(exist_ok=True, parents=True)
heroes = heroes.glob("*.png")
heroes = [(p.stem.split("_")[0], Image.open(p)) for p in heroes]

items = Path("items")
items.mkdir(exist_ok=True, parents=True)
items = items.glob("*.png")
items = [(p.stem.split("_")[0], Image.open(p)) for p in items]

stars = Path("stars")
stars.mkdir(exist_ok=True, parents=True)
stars = stars.glob("*.png")
stars = [(p.stem.split("_")[0], Image.open(p)) for p in stars]


unknown_dir = Path("unknown")
unknown_dir.mkdir(exist_ok=True, parents=True)
unknown_ranks_dir = unknown_dir / "ranks"
unknown_ranks_dir.mkdir(exist_ok=True, parents=True)
unknown_underlords_dir = unknown_dir / "underlords"
unknown_underlords_dir.mkdir(exist_ok=True, parents=True)
unknown_heroes_dir = unknown_dir / "heroes"
unknown_heroes_dir.mkdir(exist_ok=True, parents=True)
unknown_items_dir = unknown_dir / "items"
unknown_items_dir.mkdir(exist_ok=True, parents=True)
unknown_stars_dir = unknown_dir / "stars"
unknown_stars_dir.mkdir(exist_ok=True, parents=True)

out_dir = Path("processed")
out_dir.mkdir(exist_ok=True, parents=True)


class Alliance(Enum):
    assassin = auto()
    brawny = auto()
    brute = auto()
    champion = auto()
    demon = auto()
    dragon = auto()
    fallen = auto()
    healer = auto()
    heartless = auto()
    human = auto()
    hunter = auto()
    knight = auto()
    mage = auto()
    magus = auto()
    poisoner = auto()
    rogue = auto()
    savage = auto()
    scaled = auto()
    shaman = auto()
    spirit = auto()
    summoner = auto()
    swordsman = auto()
    troll = auto()
    vigilant = auto()
    void = auto()
    warrior = auto()


alliance_levels = {
    Alliance.assassin: [3, 6],
    Alliance.brawny: [2, 4],
    Alliance.brute: [2, 4],
    Alliance.champion: [1],
    Alliance.demon: [1],
    Alliance.dragon: [2],
    Alliance.fallen: [3, 6],
    Alliance.healer: [2, 4],
    Alliance.heartless: [2, 4, 6],
    Alliance.human: [2, 4, 6],
    Alliance.hunter: [3, 6],
    Alliance.knight: [2, 4, 6],
    Alliance.mage: [3, 6],
    Alliance.magus: [1],
    Alliance.poisoner: [3],
    Alliance.rogue: [2, 4],
    Alliance.savage: [2, 4, 6],
    Alliance.scaled: [2, 4],
    Alliance.shaman: [2, 4, 6],
    Alliance.spirit: [3],
    Alliance.summoner: [2, 4],
    Alliance.swordsman: [3, 6],
    Alliance.troll: [2, 4],
    Alliance.vigilant: [2, 4],
    Alliance.void: [3],
    Alliance.warrior: [3, 6],
}

hero_alliances = {
    "am": [Alliance.rogue, Alliance.hunter],
    "bat": [Alliance.troll, Alliance.knight],
    "bh": [Alliance.rogue, Alliance.assassin],
    "cm": [Alliance.human, Alliance.mage],
    "dazzle": [Alliance.troll, Alliance.healer, Alliance.poisoner],
    "drow": [Alliance.heartless, Alliance.vigilant, Alliance.hunter],
    "ench": [Alliance.shaman, Alliance.healer],
    "lich": [Alliance.fallen, Alliance.mage],
    "magnus": [Alliance.savage, Alliance.shaman],
    "pa": [Alliance.assassin, Alliance.rogue],
    "sd": [Alliance.heartless, Alliance.demon],
    "slardar": [Alliance.scaled, Alliance.warrior],
    "snapfire": [Alliance.brawny, Alliance.dragon],
    "tusk": [Alliance.savage, Alliance.warrior],
    "vs": [Alliance.fallen, Alliance.heartless],
    "veno": [Alliance.scaled, Alliance.summoner, Alliance.poisoner],
    # T2
    "bristle": [Alliance.brawny, Alliance.savage],
    "ck": [Alliance.demon, Alliance.knight],
    "earth": [Alliance.spirit, Alliance.warrior],
    "jugg": [Alliance.brawny, Alliance.swordsman],
    "kunkka": [Alliance.human, Alliance.warrior, Alliance.swordsman],
    "legion": [Alliance.human, Alliance.champion],
    "luna": [Alliance.vigilant, Alliance.knight],
    "meepo": [Alliance.rogue, Alliance.summoner],
    "nature": [Alliance.shaman, Alliance.summoner],
    "pudge": [Alliance.heartless, Alliance.warrior],
    "qop": [Alliance.demon, Alliance.assassin, Alliance.poisoner],
    "spirit": [Alliance.savage, Alliance.brute],
    "storm": [Alliance.spirit, Alliance.mage],
    "wr": [Alliance.vigilant, Alliance.hunter],
    # T3
    "aba": [Alliance.fallen, Alliance.knight],
    "alch": [Alliance.brute, Alliance.rogue, Alliance.poisoner],
    "bm": [Alliance.brawny, Alliance.hunter, Alliance.shaman],
    "ember": [Alliance.spirit, Alliance.assassin, Alliance.swordsman],
    "lifestealer": [Alliance.heartless, Alliance.brute],
    "lycan": [Alliance.human, Alliance.savage, Alliance.summoner],
    "omni": [Alliance.human, Alliance.knight, Alliance.healer],
    "puck": [Alliance.dragon, Alliance.mage],
    "shaman": [Alliance.troll, Alliance.summoner],
    "slark": [Alliance.scaled, Alliance.assassin],
    "spectre": [Alliance.void, Alliance.demon],
    "terror": [Alliance.demon, Alliance.hunter, Alliance.fallen],
    "tree": [Alliance.shaman, Alliance.healer],
    # T4
    "dp": [Alliance.fallen, Alliance.heartless],
    "doom": [Alliance.demon, Alliance.brute],
    "lina": [Alliance.human, Alliance.mage],
    "druid": [Alliance.savage, Alliance.shaman, Alliance.summoner],
    "mirana": [Alliance.vigilant, Alliance.hunter],
    "pango": [Alliance.savage, Alliance.swordsman],
    "rubick": [Alliance.mage, Alliance.magus],
    "sven": [Alliance.human, Alliance.knight, Alliance.swordsman],
    "ta": [Alliance.vigilant, Alliance.void, Alliance.assassin],
    "tide": [Alliance.scaled, Alliance.warrior],
    "viper": [Alliance.dragon, Alliance.poisoner],
    "void": [Alliance.void, Alliance.spirit],
    # T5
    "axe": [Alliance.brawny, Alliance.brute],
    "dk": [Alliance.human, Alliance.dragon, Alliance.knight],
    "faceless": [Alliance.void, Alliance.assassin],
    "kotl": [Alliance.human, Alliance.mage],
    "dusa": [Alliance.scaled, Alliance.hunter],
    "troll": [Alliance.troll, Alliance.warrior],
    "wk": [Alliance.fallen, Alliance.swordsman],
}

hero_full_names = {
    "am": "Anti-Mage",
    "bat": "Batrider",
    "bh": "Bounty Hunter",
    "cm": "Crystal Maiden",
    "dazzle": "Dazzle",
    "drow": "Drow Ranger",
    "ench": "Enchantress",
    "lich": "Lich",
    "magnus": "Magnus",
    "pa": "Phantom Assassin",
    "sd": "Shadow Demon",
    "slardar": "Slardar",
    "snapfire": "Snapfire",
    "tusk": "Tusk",
    "vs": "Vengeful Spirit",
    "veno": "Venomancer",
    # T2
    "bristle": "Bristleback",
    "ck": "Chaos Knight",
    "earth": "Earth Spirit",
    "jugg": "Juggernaut",
    "kunkka": "Kunkka",
    "legion": "Legion Commander",
    "luna": "Luna",
    "meepo": "Meepo",
    "nature": "Nature's Prophet",
    "pudge": "Pudge",
    "qop": "Queen of Pain",
    "spirit": "Spirit Breaker",
    "storm": "Storm Spirit",
    "wr": "Windranger",
    # T3
    "aba": "Abaddon",
    "alch": "Alchemist",
    "bm": "Beastmaster",
    "ember": "Ember Spirit",
    "lifestealer": "Lifestealer",
    "lycan": "Lycan",
    "omni": "Omniknight",
    "puck": "Puck",
    "shaman": "Shadow Shaman",
    "slark": "Slark",
    "spectre": "Spectre",
    "terror": "Terrorblade",
    "tree": "Treant Protector",
    # T4
    "dp": "Death Prophet",
    "doom": "Doom",
    "lina": "Lina",
    "druid": "Lone Druid",
    "mirana": "Mirana",
    "pango": "Pangolier",
    "rubick": "Rubick",
    "sven": "Sven",
    "ta": "Templar Assassin",
    "tide": "Tidehunter",
    "viper": "Viper",
    "void": "Void Spirit",
    # T5
    "axe": "Axe",
    "dk": "Dragon Knight",
    "faceless": "Faceless Void",
    "kotl": "Keeper of the Light",
    "dusa": "Medusa",
    "troll": "Troll Warlord",
    "wk": "Wraith King",
}

item_full_names = {
    "antlers": "Crown of Antlers",
    "armlet": "Armlet of Mordiggian",
    "basher": "Skull Basher",
    "bf": "Battle Fury",
    "bkb": "Black King Bar",
    "blademail": "Blade Mail",
    "bloodthorn": "Bloodthorn",
    "boots": "Arcane Boots",
    "butterfly": "Butterfly",
    "chainmail": "Chainmail",
    "claymore": "Claymore",
    "craggy": "Craggy Coat",
    "dagger": "Blink Dagger",
    "dagon": "Dagon",
    "desolator": "Desolator",
    "diffusal": "Diffusal Blade",
    "eul": "Eul's Scepter",
    "fedora": "Leg Breaker's Fedora",
    "gloves": "Gloves of Haste",
    "halberd": "Heaven's Halberd",
    "headdress": "Headdress",
    "hood": "Hood of Defiance",
    "horn": "Horn of the Alpha",
    "kaden": "Kaden's Blade",
    "kaya": "Kaya",
    "lance": "Dragon Lance",
    "maelstrom": "Maelstrom",
    "mekansm": "Mekansm",
    "mkb": "Monkey King Bar",
    "mom": "Mask of Madness",
    "moonshard": "Moon Shard",
    "morbidmask": "Morbid Mask",
    "necro": "Necronomicon",
    "octarine": "Octarine Essence",
    "orb": "Refresher Orb",
    "paladin": "Paladin Sword",
    "pike": "Stonehall Pike",
    "pipe": "Pipe of Insight",
    "pirate": "Pirate Hat",
    "qb": "Quelling Blade",
    "radiance": "Radiance",
    "rapier": "Divine Rapier",
    "ristul": "Ristul Circlet",
    "satanic": "Satanic",
    "scythe": "Scythe of Vyse",
    "shako": "Witless Shako",
    "shiva": "Shiva's Guard",
    "silveredge": "Silver Edge",
    "skadi": "Eye of Skadi",
    "stonehall": "Stonehall Cloak",
    "talisman": "Talisman of Evasion",
    "tarrasque": "Heart of Tarrasque",
    "vanguard": "Vanguard",
    "vb": "Vitality Booster",
    "vesture": "Vesture of the Tyrant",
    "vlad": "Vladmir's Offering",
    "void": "Void Stone",
}

underlord_full_names = {
    "anessix": "Anessix",
    "anessix-healer": "Anessix (Healing)",
    "anessix-damage": "Anessix (Damage)",
    "enno": "Enno",
    "enno-healer": "Enno (Healing)",
    "enno-damage": "Enno (Damage)",
    "hobgen": "Hobgen",
    "hobgen-support": "Hobgen (Support)",
    "hobgen-damage": "Hobgen (Damage)",
    "jull": "Jull",
    "jull-healer": "Jull (Healing)",
    "jull-damage": "Jull (Damage)",
    None: "",
}

item_alliances = {
    "antlers": Alliance.hunter,
    "fedora": Alliance.brute,
    "pirate": Alliance.swordsman,
    "ristul": Alliance.demon,
}


@dataclass
class Hero:
    name: str
    stars: int
    item: Optional[str] = None

    def as_dict(self):
        return {
            "name": hero_full_names[self.name],
            "stars": self.stars,
            "item": item_full_names[self.item] if self.item else None,
        }


@dataclass
class AllianceCounter:
    name: Alliance
    total: int = 0

    @property
    def active(self):
        for level in sorted(alliance_levels[self.name], reverse=True):
            if self.total >= level:
                return level
        return 0


@dataclass
class ScoreboardRow:
    rank: int
    underlord: Optional[str] = None
    heroes: list[Hero] = field(default_factory=list)
    alliances: dict[Alliance, AllianceCounter] = field(default_factory=dict)

    def as_dict(self):
        return {
            "rank": self.rank,
            "underlord": underlord_full_names[self.underlord],
            "heroes": [hero.as_dict() for hero in self.heroes],
            "alliances": {
                name.name.capitalize(): {
                    "total": counter.total,
                    "active": counter.active,
                }
                for name, counter in self.alliances.items()
            },
        }


@dataclass
class Scoreboard:
    fox_rank: int = 1
    rows: list[ScoreboardRow] = field(default_factory=list)

    def as_dict(self):
        return {
            "rows": [row.as_dict() for row in self.rows],
            "fox_rank": self.fox_rank,
        }


currently_processing = None


def save_unknown(dir: Path, img: Image):
    for unk in dir.glob("*.png"):
        with Image.open(unk) as i:
            if ImageChops.difference(img, i).getbbox() is None:
                return
    i = 0
    while True:
        if not (dir / f"{i}.png").exists():
            break
        i += 1
    img.save(str(dir / f"{i}.png"))
    print(f"No match in {currently_processing}, saved {dir / f'{i}.png'}")


def get_match(src: Image, matches_list, unknown_dir: Path):
    for val, img in matches_list:
        if ImageChops.difference(src, img).getbbox() is None:
            return val
    save_unknown(unknown_dir, src)
    return None


# Portrait spacing is 60x60
# Underlord is 87 from right edge 12 from top, 40x40
#
# 1. Grab underlord
# 2. 1px below edge of underlord, scan left until first different-colored pixel
# 3. scan right until first all-blank column
# 4. scan left until first all-blank column
# 5. halfway between the two columns is the center of the last hero
# 6. grab 38x5 pixels centered on that point and starting 1px below the
#    underlord - that's the star level
# 7. grab 20x40 pixels touching the left side of the centerline and starting
#    from the top of the row - that's the hero
# 8. grab 20x20 pixels from centerline + 20px and top of row - that's the item


def process_scoreboard(scoreboard: Image.Image):
    result = Scoreboard()
    for row in range(8):
        row_obj = ScoreboardRow(rank=row + 1)
        row_top = 2 + row * 60
        underlord_top = 10 + row_top
        underlord_left = scoreboard.width - 87
        scan_row = underlord_top + 22
        for col in range(scoreboard.width - 64, scoreboard.width - 128, -1):
            if scoreboard.getpixel((col, scan_row)) == (0, 0, 0):
                underlord_left = col + 15
                underlord = scoreboard.crop(
                    (
                        underlord_left,
                        underlord_top,
                        underlord_left + 40,
                        underlord_top + 40,
                    )
                )
                row_obj.underlord = get_match(
                    underlord, underlords, unknown_underlords_dir
                )
                break
        else:
            row_obj.underlord = None

        underlord_bottom = underlord_top + 41
        search_right_edge = underlord_left - 60
        bg_pixel = scoreboard.getpixel((underlord_left - 60, underlord_bottom))
        for bench_pos in range(1, 11):
            for col in range(search_right_edge, 0, -1):
                if scoreboard.getpixel((col, underlord_bottom)) != bg_pixel:
                    rightmost_star = col
                    break
            else:
                break
            for col2 in range(max(0, rightmost_star - 42), rightmost_star):
                if scoreboard.getpixel((col2, underlord_bottom)) != bg_pixel:
                    leftmost_star = col2 + 1
                    break
            else:
                break
            center = (leftmost_star + rightmost_star) // 2
            star_value = scoreboard.crop(
                (
                    leftmost_star + 7,
                    underlord_bottom,
                    leftmost_star + 8,
                    underlord_bottom + 1,
                )
            )
            try:
                star_value = int(get_match(star_value, stars, unknown_stars_dir))
            except TypeError:
                print(f"Error processing {currently_processing}")
                print(
                    f"Star value: ",
                    (
                        leftmost_star + 7,
                        underlord_bottom,
                        leftmost_star + 9,
                        underlord_bottom + 2,
                    ),
                )
                scoreboard.crop(
                    (
                        leftmost_star,
                        row_top,
                        rightmost_star,
                        row_top + 60,
                    )
                ).save("error.png")
                raise

            hero = scoreboard.crop(
                (
                    center - 6,
                    row_top + 26,
                    center,
                    row_top + 34,
                )
            )
            hero = get_match(hero, heroes, unknown_heroes_dir)
            if not hero:
                print(
                    f"Unknown hero in {currently_processing}, row "
                    f"{row + 1}, bench position {bench_pos}\n"
                )

            for h in row_obj.heroes:
                if h.name == hero:
                    h.stars = max(h.stars, star_value)
                    hero = h
                    break
            else:
                for alliance in hero_alliances[hero]:
                    row_obj.alliances.setdefault(
                        alliance, AllianceCounter(alliance)
                    ).total += 1
                hero = Hero(name=hero, stars=star_value)
                row_obj.heroes.append(hero)

            item = scoreboard.crop(
                (
                    center + 20,
                    row_top + 7,
                    center + 27,
                    row_top + 20,
                )
            )
            item = get_match(item, items, unknown_items_dir)
            if item in item_alliances:
                if item_alliances[item] in row_obj.alliances:
                    if item_alliances[item] not in hero_alliances[hero.name]:
                        row_obj.alliances[item_alliances[item]].total += 1
            hero.item = item
            search_right_edge = center - 36

        result.rows.append(row_obj)

    assert len(result.rows) == 8
    return result


for prefix in prefixes:
    print("Processing", prefix)
    currently_processing = prefix
    rank_img = sources / f"{prefix}_fox_rank.png"
    with Image.open(rank_img) as i:
        rank = get_match(i, ranks, unknown_ranks_dir)
    scoreboard_img = sources / f"{prefix}_teams.png"
    with Image.open(scoreboard_img) as i:
        scoreboard = process_scoreboard(i)
        scoreboard.fox_rank = rank
    assert list(unknown_dir.glob("**/*.png")) == []
    with (out_dir / f"{prefix}.json").open("w") as f:
        json.dump(scoreboard.as_dict(), f, indent=2)
