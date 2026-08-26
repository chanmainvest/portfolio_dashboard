# Runtime image for the portfolio report builder — used by the nightly
# Jenkins pipeline.
#
# The Jenkins container has no Python/uv, so the pipeline does
# `docker compose run --rm report ...` against this image. The repo root is
# bind-mounted at /app (input workbook, generated HTML/JSON and the SQLite
# market-data cache all live on the host, shared with the local
# `uv run python build_portfolio_report.py` workflow — no second cache).
#
# Build:  docker compose build report
# Run:    docker compose run --rm report
#
# No secrets are needed (public yfinance data only).

FROM python:3.12-slim

# curl: uv bootstrap. That's all the system deps the report needs.
RUN apt-get update && apt-get install -y --no-install-recommends curl \
    && rm -rf /var/lib/apt/lists/*

# ---- uv (matches the host toolchain) ----------------------------------------
# Pinned for reproducibility; update alongside the host `uv` if needed.
ARG UV_VERSION=0.5.11
RUN curl -LsSf https://astral.sh/uv/${UV_VERSION}/install.sh | sh \
    && mv /root/.local/bin/uv /usr/local/bin/uv \
    && mv /root/.local/bin/uvx /usr/local/bin/uvx

# Dependencies go to /opt/venv, NOT /app/.venv: the Jenkins run bind-mounts
# the live repo over /app, which would shadow a baked-in venv and make the
# image lose its Python environment at runtime.
ENV UV_LINK_MODE=copy \
    UV_COMPILE_BYTECODE=1 \
    UV_PYTHON_DOWNLOADS=never \
    UV_PROJECT_ENVIRONMENT=/opt/venv

WORKDIR /app

# ---- Install dependencies (cached unless pyproject/lock change) -------------
# Copy only the manifest first so dependency resolution is layer-cached.
# The build script itself is NOT copied: it comes from the /app bind mount,
# so workspace edits take effect on the next run without an image rebuild.
COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-dev

ENV PATH="/opt/venv/bin:$PATH" \
    PYTHONUNBUFFERED=1

# Default: regenerate the grandma report (override via
# `docker compose run --rm report <cmd...>`).
CMD ["python", "build_portfolio_report.py", "--input", "grandma_investment_portfolio.xlsx", "--output", "grandma.html"]
