from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
import unicodedata

import yaml


DEFAULT_TECH_ALIASES: dict[str, list[str]] = {
    "Playwright": ["playwright"],
    "Cypress": ["cypress"],
    "Selenium": ["selenium", "selenium webdriver"],
    "Python": ["python", "pytest", "py.test"],
    "JavaScript": ["javascript", "js"],
    "TypeScript": ["typescript", "ts"],
    "Java": ["java"],
    "C#": ["c#", "c sharp", "csharp"],
    "API testing": ["api testing", "api tests", "api test", "rest api", "rest", "postman"],
    "SQL": ["sql", "database", "databases"],
    "WCAG": ["wcag", "accessibility", "a11y"],
    "Allure": ["allure", "allure report", "allure reports"],
    "Git": ["git", "github", "gitlab", "bitbucket"],
    "CI/CD": ["ci/cd", "cicd", "continuous integration", "github actions", "jenkins"],
    "Docker": ["docker", "container", "containers"],
    "Jira": ["jira", "confluence"],
    "TestRail": ["testrail", "test rail"],
    "Robot Framework": ["robot framework"],
    "Cucumber": ["cucumber", "gherkin", "bdd"],
    "Manual testing": ["manual testing", "manual tests", "test cases"],
    "Test automation": ["test automation", "automated tests", "automation testing"],
    "Regression testing": ["regression testing", "regression tests"],
    "Performance testing": ["performance testing", "load testing", "jmeter"],
    "Security testing": ["security testing", "penetration testing"],
    "Mobile testing": ["mobile testing", "android", "ios"],
    "Agile": ["agile", "scrum", "kanban"],
    "AI coding assistants": ["ai coding", "ai assistant", "copilot", "chatgpt"],
}


class CvProfileError(RuntimeError):
    pass


class CvProfileNotFoundError(CvProfileError):
    pass


@dataclass(frozen=True)
class CandidateProfile:
    skills: list[str]
    projects: list[str]
    roles: list[str]


@dataclass(frozen=True)
class RequirementMatch:
    requirement: str
    has_skill: bool
    evidence: str
    missing_skill: str
    comment: str


@dataclass(frozen=True)
class CvMatchResult:
    matched_skills: list[str]
    missing_skills: list[str]
    match_score: int
    priority: str
    short_reason: str
    requirements: list[RequirementMatch]


def load_candidate_profile(profile_path: Path) -> CandidateProfile:
    profile_path = Path(profile_path)
    if not profile_path.exists():
        raise CvProfileNotFoundError(
            "Brak pliku profilu CV. Skopiuj data/sample_cv_profile.yml do "
            "data/private/cv_profile.yml i uzupełnij własne dane."
        )

    with profile_path.open("r", encoding="utf-8") as file:
        payload = yaml.safe_load(file) or {}

    if not isinstance(payload, dict):
        raise CvProfileError("Profil CV musi być obiektem YAML z listami skills/projects/roles.")

    return CandidateProfile(
        skills=_string_list(payload.get("skills")),
        projects=_string_list(payload.get("projects")),
        roles=_string_list(payload.get("roles")),
    )


def match_offer_to_profile(offer_text: str, profile: CandidateProfile) -> CvMatchResult:
    extracted_requirements = extract_requirements(offer_text, profile.skills)
    if not extracted_requirements:
        requirement = RequirementMatch(
            requirement="Nie wykryto technologii ani wymagań technicznych w tekście oferty",
            has_skill=False,
            evidence="",
            missing_skill="Do ręcznej analizy",
            comment=(
                "Niski confidence: oferta nie zawiera rozpoznanych technologii. "
                "Matcher nie zgaduje dopasowania."
            ),
        )
        return CvMatchResult(
            matched_skills=[],
            missing_skills=["Do ręcznej analizy"],
            match_score=0,
            priority="LOW",
            short_reason="Nie wykryto technologii w ofercie. Wymagana ręczna analiza.",
            requirements=[requirement],
        )

    matched_skills: list[str] = []
    missing_skills: list[str] = []
    requirement_matches: list[RequirementMatch] = []

    for requirement in extracted_requirements:
        profile_skill = _matching_profile_skill(requirement, profile.skills)
        if profile_skill:
            matched_skills.append(profile_skill)
            requirement_matches.append(
                RequirementMatch(
                    requirement=requirement,
                    has_skill=True,
                    evidence=profile_skill,
                    missing_skill="",
                    comment=f"Wymaganie znalezione w profilu jako: {profile_skill}.",
                )
            )
        else:
            missing_skills.append(requirement)
            requirement_matches.append(
                RequirementMatch(
                    requirement=requirement,
                    has_skill=False,
                    evidence="",
                    missing_skill=requirement,
                    comment="Wymaganie występuje w ofercie, ale nie ma go na liście skills.",
                )
            )

    matched_unique = _unique_keep_order(matched_skills)
    missing_unique = _unique_keep_order(missing_skills)
    score = round((len(matched_unique) / len(extracted_requirements)) * 100)
    priority = priority_from_score(score)
    short_reason = _short_reason(score, matched_unique, missing_unique)

    return CvMatchResult(
        matched_skills=matched_unique,
        missing_skills=missing_unique,
        match_score=score,
        priority=priority,
        short_reason=short_reason,
        requirements=requirement_matches,
    )


def extract_requirements(offer_text: str, profile_skills: list[str]) -> list[str]:
    normalized_text = _normalize_text(offer_text)
    if not normalized_text:
        return []

    requirements: list[str] = []
    for canonical, aliases in _requirement_catalog(profile_skills).items():
        if any(_contains_alias(normalized_text, alias) for alias in aliases):
            requirements.append(canonical)

    return sorted(_unique_keep_order(requirements), key=str.lower)


def priority_from_score(score: int) -> str:
    if score >= 75:
        return "HIGH"
    if score >= 50:
        return "MEDIUM"
    return "LOW"


def _requirement_catalog(profile_skills: list[str]) -> dict[str, list[str]]:
    catalog = {skill: aliases[:] for skill, aliases in DEFAULT_TECH_ALIASES.items()}
    for skill in profile_skills:
        if skill not in catalog:
            catalog[skill] = [skill]
        elif skill not in catalog[skill]:
            catalog[skill].append(skill)
    return catalog


def _matching_profile_skill(requirement: str, profile_skills: list[str]) -> str:
    requirement_aliases = [_normalize_text(requirement)]
    requirement_aliases.extend(
        _normalize_text(alias) for alias in DEFAULT_TECH_ALIASES.get(requirement, [])
    )

    for skill in profile_skills:
        skill_aliases = [_normalize_text(skill)]
        skill_aliases.extend(_normalize_text(alias) for alias in DEFAULT_TECH_ALIASES.get(skill, []))
        if set(requirement_aliases) & set(skill_aliases):
            return skill

    return ""


def _contains_alias(normalized_text: str, alias: str) -> bool:
    normalized_alias = _normalize_text(alias)
    if not normalized_alias:
        return False

    pattern = rf"(?<![a-z0-9]){re.escape(normalized_alias)}(?![a-z0-9])"
    return re.search(pattern, normalized_text) is not None


def _normalize_text(text: str) -> str:
    ascii_text = unicodedata.normalize("NFKD", str(text).casefold())
    ascii_text = "".join(char for char in ascii_text if not unicodedata.combining(char))
    ascii_text = ascii_text.replace("ł", "l")
    ascii_text = re.sub(r"[^a-z0-9#+./-]+", " ", ascii_text)
    return re.sub(r"\s+", " ", ascii_text).strip()


def _string_list(value: object) -> list[str]:
    if value is None:
        return []
    if not isinstance(value, list):
        raise CvProfileError("Pola skills, projects i roles muszą być listami.")

    return [str(item).strip() for item in value if str(item).strip()]


def _unique_keep_order(values: list[str]) -> list[str]:
    seen: set[str] = set()
    unique_values: list[str] = []
    for value in values:
        key = _normalize_text(value)
        if key in seen:
            continue
        seen.add(key)
        unique_values.append(value)
    return unique_values


def _short_reason(score: int, matched_skills: list[str], missing_skills: list[str]) -> str:
    matched_count = len(matched_skills)
    total_count = matched_count + len(missing_skills)
    if not missing_skills:
        return f"Dopasowanie {score}%: profil pokrywa wszystkie wykryte wymagania ({matched_count}/{total_count})."

    missing_preview = "; ".join(missing_skills[:5])
    return (
        f"Dopasowanie {score}%: profil pokrywa {matched_count}/{total_count} "
        f"wykrytych wymagań. Braki: {missing_preview}."
    )
