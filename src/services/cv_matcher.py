from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
import unicodedata

import yaml


DEFAULT_TECH_ALIASES: dict[str, list[str]] = {
    "Playwright": ["playwright"],
    "Python": ["python"],
    "Pytest": ["pytest", "py.test"],
    "Test automation": [
        "test automation",
        "automated tests",
        "automation testing",
        "automated testing",
        "test automatyzacja",
        "testy automatyczne",
        "automatyzacja testów",
        "automatyzacja testow",
    ],
    "Page Object Model": ["page object model", "page object pattern", "pom"],
    "Manual testing": ["manual testing", "manual tests", "manual tester"],
    "Functional testing": ["functional testing", "functional tests"],
    "Regression testing": ["regression testing", "regression tests"],
    "Test case design": ["test case design", "test design", "test cases", "test scenarios"],
    "REST API": ["rest api", "api testing", "api tests", "api test", "rest", "microservices"],
    "Postman": ["postman"],
    "Swagger": ["swagger", "openapi", "open api"],
    "SQL": ["sql", "database", "databases"],
    "Test data preparation": ["test data", "test data preparation", "data preparation"],
    "Git": ["git", "github", "gitlab", "bitbucket"],
    "Jira": ["jira", "confluence"],
    "Allure": ["allure", "allure report", "allure reports"],
    "CI/CD": ["ci/cd", "cicd", "continuous integration", "github actions", "jenkins"],
    "Agile": ["agile", "scrum", "kanban"],
    "WCAG": ["wcag"],
    "Accessibility testing": ["accessibility", "a11y", "accessibility testing"],
    "AI-assisted testing": [
        "ai-assisted testing",
        "ai assisted testing",
        "ai testing",
        "ai in testing",
    ],
    "Quality Engineering": ["quality engineering", "quality engineer"],
    "Cypress": ["cypress"],
    "Selenium": ["selenium", "selenium webdriver"],
    "Java": ["java"],
    "C#": ["c#", "c sharp", "csharp"],
    "JavaScript": ["javascript", "js"],
    "TypeScript": ["typescript", "ts"],
    "Docker": ["docker", "container", "containers"],
    "TestRail": ["testrail", "test rail"],
    "XRay": ["xray", "x-ray"],
    "Robot Framework": ["robot framework"],
    "Cucumber": ["cucumber", "gherkin", "bdd"],
    "Performance testing": ["performance testing", "load testing", "jmeter"],
    "Security testing": ["security testing", "penetration testing"],
    "Mobile testing": ["mobile testing", "android", "ios"],
    "Ferryt": ["ferryt"],
}

REQUIREMENT_CATEGORIES: dict[str, str] = {
    "Playwright": "core_skills",
    "Python": "core_skills",
    "Pytest": "core_skills",
    "Test automation": "core_skills",
    "Page Object Model": "core_skills",
    "Manual testing": "qa_skills",
    "Functional testing": "qa_skills",
    "Regression testing": "qa_skills",
    "Test case design": "qa_skills",
    "REST API": "api_data",
    "Postman": "api_data",
    "Swagger": "api_data",
    "SQL": "api_data",
    "Test data preparation": "api_data",
    "Git": "tools_workflow",
    "Jira": "tools_workflow",
    "Allure": "tools_workflow",
    "CI/CD": "tools_workflow",
    "Agile": "tools_workflow",
    "WCAG": "specialization",
    "Accessibility testing": "specialization",
    "AI-assisted testing": "specialization",
    "Quality Engineering": "specialization",
    "Cypress": "primary_stack",
    "Java": "primary_stack",
    "C#": "primary_stack",
}

CATEGORY_WEIGHTS = {
    "core_skills": 3.0,
    "qa_skills": 2.0,
    "api_data": 2.0,
    "tools_workflow": 1.0,
    "specialization": 2.0,
    "primary_stack": 2.5,
    "other": 1.5,
}
LOW_SIGNAL_REQUIREMENTS = {"Git", "Jira", "Agile"}
PRIMARY_PROFILE_STACK = {"Playwright", "Python", "Pytest"}
PRIMARY_STACK_MISMATCH_REQUIREMENTS = {"Java", "C#", "Cypress"}
NICHE_DOMAIN_TOOLS = {"Ferryt"}
CORE_SKILL_MISSING_PENALTY = 12


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
    category: str
    weight: float
    has_skill: bool
    evidence: str
    missing_skill: str
    comment: str


@dataclass(frozen=True)
class CvMatchResult:
    matched_skills: list[str]
    missing_skills: list[str]
    match_score: int
    confidence_score: int
    priority: str
    short_reason: str
    requirements: list[RequirementMatch]
    penalties: list[str]
    raw_score: int


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
    confidence_score = confidence_from_requirements(extracted_requirements)
    if not extracted_requirements:
        requirement = RequirementMatch(
            requirement="Nie wykryto technologii ani wymagań technicznych w tekście oferty",
            category="other",
            weight=0.0,
            has_skill=False,
            evidence="",
            missing_skill="Do ręcznej analizy",
            comment=(
                "Confidence 0%: oferta nie zawiera rozpoznanych wymagań. "
                "Matcher nie zgaduje dopasowania."
            ),
        )
        return CvMatchResult(
            matched_skills=[],
            missing_skills=["Do ręcznej analizy"],
            match_score=0,
            confidence_score=0,
            priority="LOW",
            short_reason="Dopasowanie 0%, confidence 0%: nie wykryto wymagań w ofercie.",
            requirements=[requirement],
            penalties=[],
            raw_score=0,
        )

    normalized_text = _normalize_text(offer_text)
    matched_skills: list[str] = []
    missing_skills: list[str] = []
    requirement_matches: list[RequirementMatch] = []
    total_weight = 0.0
    matched_weight = 0.0

    for requirement in extracted_requirements:
        category = _requirement_category(requirement)
        weight = _requirement_weight(requirement)
        total_weight += weight
        profile_skill = _matching_profile_skill(requirement, profile.skills)
        if profile_skill:
            matched_weight += weight
            matched_skills.append(profile_skill)
            requirement_matches.append(
                RequirementMatch(
                    requirement=requirement,
                    category=category,
                    weight=weight,
                    has_skill=True,
                    evidence=profile_skill,
                    missing_skill="",
                    comment=(
                        f"Kategoria: {category}, waga: {weight:g}. "
                        f"Wymaganie znalezione w profilu jako: {profile_skill}."
                    ),
                )
            )
        else:
            missing_skills.append(requirement)
            requirement_matches.append(
                RequirementMatch(
                    requirement=requirement,
                    category=category,
                    weight=weight,
                    has_skill=False,
                    evidence="",
                    missing_skill=requirement,
                    comment=(
                        f"Kategoria: {category}, waga: {weight:g}. "
                        "Wymaganie występuje w ofercie, ale nie ma go na liście skills."
                    ),
                )
            )

    raw_score = round((matched_weight / total_weight) * 100) if total_weight else 0
    penalty_points, penalties = _score_penalties(
        extracted_requirements,
        missing_skills,
        normalized_text,
        profile,
    )
    score_after_penalties = max(0, raw_score - penalty_points)
    capped_score = min(score_after_penalties, _max_score_for_requirement_count(len(extracted_requirements)))
    priority = priority_from_score(capped_score, confidence_score)
    matched_unique = _unique_keep_order(matched_skills)
    missing_unique = _unique_keep_order(missing_skills)

    return CvMatchResult(
        matched_skills=matched_unique,
        missing_skills=missing_unique,
        match_score=capped_score,
        confidence_score=confidence_score,
        priority=priority,
        short_reason=_short_reason(
            score=capped_score,
            confidence_score=confidence_score,
            raw_score=raw_score,
            penalty_points=penalty_points,
            requirement_count=len(extracted_requirements),
            matched_skills=matched_unique,
            missing_skills=missing_unique,
            penalties=penalties,
        ),
        requirements=requirement_matches,
        penalties=penalties,
        raw_score=raw_score,
    )


def extract_requirements(offer_text: str, profile_skills: list[str]) -> list[str]:
    normalized_text = _normalize_text(offer_text)
    if not normalized_text:
        return []

    requirements: list[str] = []
    for canonical, aliases in _requirement_catalog(profile_skills).items():
        if any(_contains_alias(normalized_text, alias) for alias in aliases):
            requirements.append(canonical)

    return _unique_keep_order(requirements)


def confidence_from_requirements(requirements: list[str]) -> int:
    count = len(requirements)
    if count == 0:
        return 0

    if count == 1:
        score = 25
    elif count == 2:
        score = 40
    elif count == 3:
        score = 55
    elif count == 4:
        score = 70
    elif count == 5:
        score = 80
    else:
        score = 90

    categories = {_requirement_category(requirement) for requirement in requirements}
    score += min(10, max(0, len(categories) - 1) * 3)

    if categories == {"tools_workflow"}:
        score = min(score, 35)
    if count <= 2:
        score = min(score, 40)

    return _clamp_score(score)


def priority_from_score(score: int, confidence_score: int = 100) -> str:
    if score >= 75 and confidence_score >= 65:
        return "HIGH"
    if score >= 50 and confidence_score >= 45:
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


def _score_penalties(
    requirements: list[str],
    missing_skills: list[str],
    normalized_text: str,
    profile: CandidateProfile,
) -> tuple[int, list[str]]:
    penalty_points = 0
    penalties: list[str] = []

    if _is_manual_only_role(normalized_text):
        penalty_points += 20
        penalties.append("manual-only role: -20")

    missing_core = [
        skill
        for skill in missing_skills
        if _requirement_category(skill) == "core_skills"
    ]
    if missing_core:
        penalty = min(25, len(missing_core) * CORE_SKILL_MISSING_PENALTY)
        penalty_points += penalty
        penalties.append(f"brak wymaganego core skilla ({', '.join(missing_core)}): -{penalty}")

    missing_niche_tools = [
        tool
        for tool in requirements
        if tool in NICHE_DOMAIN_TOOLS and not _matching_profile_skill(tool, profile.skills)
    ]
    if missing_niche_tools:
        penalty = min(15, len(missing_niche_tools) * 10)
        penalty_points += penalty
        penalties.append(
            f"niszowe narzędzie domenowe poza profilem ({', '.join(missing_niche_tools)}): -{penalty}"
        )

    missing_primary_stack = [
        requirement
        for requirement in requirements
        if requirement in PRIMARY_STACK_MISMATCH_REQUIREMENTS
        and not _matching_profile_skill(requirement, profile.skills)
    ]
    has_profile_primary_stack_in_offer = any(
        requirement in PRIMARY_PROFILE_STACK for requirement in requirements
    )
    if missing_primary_stack and not has_profile_primary_stack_in_offer:
        penalty_points += 20
        penalties.append(
            "primary stack mismatch "
            f"({', '.join(missing_primary_stack)} zamiast Playwright/Python): -20"
        )
    elif missing_primary_stack:
        penalty_points += 10
        penalties.append(
            f"częściowy primary stack mismatch ({', '.join(missing_primary_stack)}): -10"
        )

    if _seniority_too_high(normalized_text, profile.roles):
        penalty_points += 12
        penalties.append("seniority too high: -12")

    return penalty_points, penalties


def _is_manual_only_role(normalized_text: str) -> bool:
    manual_terms = [
        "manual only",
        "manual-only",
        "manual tester",
        "manual qa",
        "manual testing",
        "testy manualne",
        "tester manualny",
    ]
    automation_terms = [
        "test automation",
        "automated test",
        "automation testing",
        "automated testing",
        "playwright",
        "pytest",
        "selenium",
        "cypress",
        "python",
        "testy automatyczne",
        "automatyzacja testow",
        "automatyzujacy",
        "tester automatyzujacy",
    ]
    return any(term in normalized_text for term in manual_terms) and not any(
        term in normalized_text for term in automation_terms
    )


def _seniority_too_high(normalized_text: str, profile_roles: list[str]) -> bool:
    senior_profile = any(
        _contains_alias(_normalize_text(role), term)
        for role in profile_roles
        for term in ["senior", "lead", "principal", "staff", "expert"]
    )
    if senior_profile:
        return False

    if re.search(r"(?<!\d)([5-9]|1\d)\+?\s*(years|lat)", normalized_text):
        return True

    senior_terms = [
        "senior",
        "lead qa",
        "principal",
        "staff engineer",
        "expert",
    ]
    return any(term in normalized_text for term in senior_terms)


def _max_score_for_requirement_count(requirement_count: int) -> int:
    if requirement_count <= 0:
        return 0
    if requirement_count <= 2:
        return 50
    if requirement_count == 3:
        return 65
    return 100


def _requirement_category(requirement: str) -> str:
    return REQUIREMENT_CATEGORIES.get(requirement, "other")


def _requirement_weight(requirement: str) -> float:
    if requirement in LOW_SIGNAL_REQUIREMENTS:
        return 0.4
    return CATEGORY_WEIGHTS[_requirement_category(requirement)]


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


def _clamp_score(score: int) -> int:
    return max(0, min(100, int(round(score))))


def _short_reason(
    score: int,
    confidence_score: int,
    raw_score: int,
    penalty_points: int,
    requirement_count: int,
    matched_skills: list[str],
    missing_skills: list[str],
    penalties: list[str],
) -> str:
    matched_count = len(matched_skills)
    total_count = matched_count + len(missing_skills)
    reason = (
        f"Dopasowanie {score}%, confidence {confidence_score}%: "
        f"profil pokrywa {matched_count}/{total_count} wykrytych wymagań "
        f"(wykryto {requirement_count}, wynik ważony przed korektami: {raw_score}%)."
    )

    if penalty_points:
        reason = f"{reason} Kary: {'; '.join(penalties)}."

    max_score = _max_score_for_requirement_count(requirement_count)
    if max_score < 100:
        reason = f"{reason} Limit za małą liczbę wymagań: max {max_score}%."

    if missing_skills:
        reason = f"{reason} Braki: {'; '.join(missing_skills[:5])}."

    return reason
