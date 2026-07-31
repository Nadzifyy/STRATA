import json
import os
from pathlib import Path

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from groq import Groq
from pydantic import BaseModel, Field


load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"

app = FastAPI(title="SWOT & TOWS Strategy Generator")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


class TowsRequest(BaseModel):
    strengths: list[str] = Field(default_factory=list)
    weaknesses: list[str] = Field(default_factory=list)
    opportunities: list[str] = Field(default_factory=list)
    threats: list[str] = Field(default_factory=list)


SYSTEM_PROMPT = (
    "You are a pure SWOT and TOWS Strategic Engine. ONLY use provided factors. "
    "Do NOT add outside knowledge. Return strictly valid JSON containing "
    "swot_analysis, so_strategies, st_strategies, wo_strategies, and "
    "wt_strategies. swot_analysis must be an object with strengths, weaknesses, "
    "opportunities, and threats arrays. Every analysis item and strategy must use "
    "factor citations like [S1], [O1]."
)


def clean_factors(items: list[str]) -> list[str]:
    return [item.strip() for item in items if item and item.strip()]


@app.get("/", include_in_schema=False)
async def index() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


@app.post("/api/generate-tows")
async def generate_tows(payload: TowsRequest) -> dict:
    factors = {
        "strengths": clean_factors(payload.strengths),
        "weaknesses": clean_factors(payload.weaknesses),
        "opportunities": clean_factors(payload.opportunities),
        "threats": clean_factors(payload.threats),
    }

    if not all(factors.values()):
        raise HTTPException(
            status_code=400,
            detail="Add at least one factor to every S, W, O, and T category.",
        )

    api_key = os.getenv("GROQ_API_KEY")
    if not api_key or api_key == "your_groq_api_key_here":
        raise HTTPException(
            status_code=500,
            detail="Set a valid GROQ_API_KEY in your .env file before generating strategies.",
        )

    formatted_factors = "\n".join(
        f"{label}:\n" + "\n".join(f"[{tag}{index}] {factor}" for index, factor in enumerate(factors[key], 1))
        for label, tag, key in (
            ("Strengths", "S", "strengths"),
            ("Weaknesses", "W", "weaknesses"),
            ("Opportunities", "O", "opportunities"),
            ("Threats", "T", "threats"),
        )
    )

    user_prompt = (
        "First, create a concise SWOT analysis: interpret each category using only "
        "the supplied factors. Then create concise, actionable TOWS strategies. "
        "Return JSON in this exact shape: {\"swot_analysis\": {\"strengths\": [], "
        "\"weaknesses\": [], \"opportunities\": [], \"threats\": []}, "
        "\"so_strategies\": [], \"st_strategies\": [], \"wo_strategies\": [], "
        "\"wt_strategies\": []}.\n\n"
        f"{formatted_factors}"
    )

    try:
        client = Groq(api_key=api_key)
        completion = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.25,
        )
        result = json.loads(completion.choices[0].message.content)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=502, detail="Groq returned invalid JSON.") from exc
    except Exception as exc:
        raise HTTPException(status_code=502, detail=f"Unable to generate strategies: {exc}") from exc

    required_keys = ("so_strategies", "st_strategies", "wo_strategies", "wt_strategies")
    swot_keys = ("strengths", "weaknesses", "opportunities", "threats")
    if not all(isinstance(result.get(key), list) for key in required_keys):
        raise HTTPException(status_code=502, detail="Groq returned an unexpected strategy format.")
    if not isinstance(result.get("swot_analysis"), dict) or not all(
        isinstance(result["swot_analysis"].get(key), list) for key in swot_keys
    ):
        raise HTTPException(status_code=502, detail="Groq returned an unexpected SWOT format.")

    return {
        "swot_analysis": {key: result["swot_analysis"][key] for key in swot_keys},
        **{key: result[key] for key in required_keys},
    }
