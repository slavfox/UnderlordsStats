import json
from collections import defaultdict
from dataclasses import asdict, dataclass
from pathlib import Path
from statistics import mean, stdev

games_dir = Path(__file__).parent.parent / "processed"
total_games = len(list(games_dir.glob("*.json")))

data = []
for game_file in games_dir.glob("*.json"):
    with game_file.open() as f:
        data.extend(json.load(f)["rows"])

hero_tiers = {
    "Anti-Mage": 1,
    "Batrider": 1,
    "Bounty Hunter": 1,
    "Crystal Maiden": 1,
    "Dazzle": 1,
    "Drow Ranger": 1,
    "Enchantress": 1,
    "Lich": 1,
    "Magnus": 1,
    "Phantom Assassin": 1,
    "Shadow Demon": 1,
    "Slardar": 1,
    "Snapfire": 1,
    "Tusk": 1,
    "Vengeful Spirit": 1,
    "Venomancer": 1,
    # T2
    "Bristleback": 2,
    "Chaos Knight": 2,
    "Earth Spirit": 2,
    "Juggernaut": 2,
    "Kunkka": 2,
    "Legion Commander": 2,
    "Luna": 2,
    "Meepo": 2,
    "Nature's Prophet": 2,
    "Pudge": 2,
    "Queen of Pain": 2,
    "Spirit Breaker": 2,
    "Storm Spirit": 2,
    "Windranger": 2,
    # T3
    "Abaddon": 3,
    "Alchemist": 3,
    "Beastmaster": 3,
    "Ember Spirit": 3,
    "Lifestealer": 3,
    "Lycan": 3,
    "Omniknight": 3,
    "Puck": 3,
    "Shadow Shaman": 3,
    "Slark": 3,
    "Spectre": 3,
    "Terrorblade": 3,
    "Treant Protector": 3,
    # T4
    "Death Prophet": 4,
    "Doom": 4,
    "Lina": 4,
    "Lone Druid": 4,
    "Mirana": 4,
    "Pangolier": 4,
    "Rubick": 4,
    "Sven": 4,
    "Templar Assassin": 4,
    "Tidehunter": 4,
    "Viper": 4,
    "Void Spirit": 4,
    # T5
    "Axe": 5,
    "Dragon Knight": 5,
    "Faceless Void": 5,
    "Keeper of the Light": 5,
    "Medusa": 5,
    "Troll Warlord": 5,
    "Wraith King": 5,
}

tier_heroes_seen_by_stars = {
    1: {1: 0, 2: 0, 3: 0, "total": 0},
    2: {1: 0, 2: 0, 3: 0, "total": 0},
    3: {1: 0, 2: 0, 3: 0, "total": 0},
    4: {1: 0, 2: 0, 3: 0, "total": 0},
    5: {1: 0, 2: 0, 3: 0, "total": 0},
}

hero_finishes = defaultdict(lambda: {1: [], 2: [], 3: []})
underlord_finishes = defaultdict(list)
item_finishes = defaultdict(list)
alliance_finishes = defaultdict(list)
active_alliance_finishes = defaultdict(list)

for row in data:
    if not row["underlord"]:
        continue
    underlord_finishes[row["underlord"]].append(row["rank"])
    underlord_finishes[row["underlord"].split(" ")[0]].append(row["rank"])
    for alliance, a in row["alliances"].items():
        if a["active"]:
            alliance_finishes[alliance].append(row["rank"])
            active_alliance_finishes[(alliance, a["active"])].append(
                row["rank"]
            )
    for hero in row["heroes"]:
        tier_heroes_seen_by_stars[hero_tiers[hero["name"]]][hero["stars"]] += 1
        tier_heroes_seen_by_stars[hero_tiers[hero["name"]]]["total"] += 1
        hero_finishes[hero["name"]][hero["stars"]].append(row["rank"])
        if hero["item"]:
            item_finishes[hero["item"]].append(row["rank"])

finish_rates_by_tier = {
    n: {
        1: tier_heroes_seen_by_stars[n][1]
        / tier_heroes_seen_by_stars[n]["total"],
        2: tier_heroes_seen_by_stars[n][2]
        / tier_heroes_seen_by_stars[n]["total"],
        3: tier_heroes_seen_by_stars[n][3]
        / tier_heroes_seen_by_stars[n]["total"],
        "total": tier_heroes_seen_by_stars[n]["total"] / len(data),
    }
    for n in range(1, 6)
}


@dataclass(slots=True)
class HeroRow:
    name: str
    total_pickrate: float
    pickrate_1: float
    pickrate_2: float
    pickrate_3: float

    avg_finish_1: float
    finish_std_dev_1: float
    avg_finish_2: float
    finish_std_dev_2: float
    avg_finish_3: float
    finish_std_dev_3: float
    avg_finish_total: float
    finish_std_dev_total: float


@dataclass(slots=True)
class UnderlordRow:
    name: str
    pickrate: float
    avg_finish: float
    finish_std_dev: float


@dataclass(slots=True)
class AllianceRow:
    name: str
    pickrate: float
    avg_finish: float
    finish_std_dev: float


@dataclass(slots=True)
class ActiveAllianceRow:
    name: str
    level: int
    pickrate: float
    avg_finish: float
    finish_std_dev: float


@dataclass(slots=True)
class ItemRow:
    name: str
    pickrate: float
    avg_finish: float
    finish_std_dev: float


@dataclass(slots=True)
class UnderlordsMeta:
    games_analyzed: int
    rows_analyzed: int
    tier_star_frequencies: dict[int, dict[int, float]]
    hero_stats: list[HeroRow]
    underlord_stats: list[UnderlordRow]
    alliance_stats: list[AllianceRow]
    active_alliance_stats: list[ActiveAllianceRow]
    item_stats: list[ItemRow]


def build_stats():
    tier_star_frequencies = finish_rates_by_tier
    hero_stats = [
        HeroRow(
            name=hero,
            total_pickrate=(
                len(hero_finishes[hero][1])
                + len(hero_finishes[hero][2])
                + len(hero_finishes[hero][3])
            )
            / len(data),
            pickrate_1=len(hero_finishes[hero][1]) / len(data),
            pickrate_2=len(hero_finishes[hero][2]) / len(data),
            pickrate_3=len(hero_finishes[hero][3]) / len(data),
            avg_finish_1=mean(hero_finishes[hero][1])
            if hero_finishes[hero][1]
            else None,
            finish_std_dev_1=stdev(hero_finishes[hero][1])
            if len(hero_finishes[hero][1]) > 1
            else None,
            avg_finish_2=mean(hero_finishes[hero][2])
            if hero_finishes[hero][2]
            else None,
            finish_std_dev_2=stdev(hero_finishes[hero][2])
            if len(hero_finishes[hero][2]) > 1
            else None,
            avg_finish_3=mean(hero_finishes[hero][3])
            if hero_finishes[hero][3]
            else None,
            finish_std_dev_3=stdev(hero_finishes[hero][3])
            if len(hero_finishes[hero][3]) > 1
            else None,
            avg_finish_total=mean(
                hero_finishes[hero][1]
                + hero_finishes[hero][2]
                + hero_finishes[hero][3]
            ),
            finish_std_dev_total=stdev(
                hero_finishes[hero][1]
                + hero_finishes[hero][2]
                + hero_finishes[hero][3]
            ),
        )
        for hero in hero_finishes
    ]
    underlord_stats = [
        UnderlordRow(
            name=underlord,
            pickrate=len(underlord_finishes[underlord]) / len(data),
            avg_finish=mean(underlord_finishes[underlord]),
            finish_std_dev=stdev(underlord_finishes[underlord])
            if len(underlord_finishes[underlord]) > 1
            else None,
        )
        for underlord in underlord_finishes
    ]
    alliance_stats = [
        AllianceRow(
            name=alliance,
            pickrate=len(alliance_finishes[alliance]) / len(data),
            avg_finish=mean(alliance_finishes[alliance]),
            finish_std_dev=stdev(alliance_finishes[alliance])
            if len(alliance_finishes[alliance]) > 1
            else None,
        )
        for alliance in alliance_finishes
    ]
    active_alliance_stats = [
        ActiveAllianceRow(
            name=alliance[0],
            level=alliance[1],
            pickrate=len(active_alliance_finishes[alliance]) / len(data),
            avg_finish=mean(active_alliance_finishes[alliance]),
            finish_std_dev=stdev(active_alliance_finishes[alliance])
            if len(active_alliance_finishes[alliance]) > 1
            else None,
        )
        for alliance in active_alliance_finishes
    ]
    item_stats = [
        ItemRow(
            name=item,
            pickrate=len(item_finishes[item]) / len(data),
            avg_finish=mean(item_finishes[item]),
            finish_std_dev=stdev(item_finishes[item])
            if len(item_finishes[item]) > 1
            else None,
        )
        for item in item_finishes
    ]
    return UnderlordsMeta(
        games_analyzed=total_games,
        rows_analyzed=len(data),
        tier_star_frequencies=tier_star_frequencies,
        hero_stats=hero_stats,
        underlord_stats=underlord_stats,
        alliance_stats=alliance_stats,
        active_alliance_stats=active_alliance_stats,
        item_stats=item_stats,
    )


print(json.dumps(asdict(build_stats()), indent=2))
