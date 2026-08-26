// =============================================================================
// Nightly portfolio report pipeline.
//
// Regenerates grandma.html (the analytics report for
// grandma_investment_portfolio.xlsx) every night, so the report always
// reflects the latest Yahoo Finance closes. Everything runs in the
// self-contained `report` Docker image (see Dockerfile / docker-compose.yml)
// because the Jenkins controller has no Python/uv.
//
// Schedule: 03:30 local time (PDT), daily - staggered 30 min after the
// knowledge-base-nightly job. At that hour the latest available close is the
// previous US trading day (US close is 13:00 PDT). For same-evening closes
// move the cron to e.g. '30 14 * * 1-5' (both in triggers below AND the
// job-level TimerTrigger in Jenkins - see AGENTS.md).
//
// Manual run: "Build Now" in Jenkins, or trigger via the REST API.
//
// Design notes (all lessons from knowledge_base/doc/jenkins-pipeline.md):
// * The repo is bind-mounted into the Jenkins container at
//   /work/chanmainvest/portfolio_dashboard (host B:\). `customWorkspace`
//   makes every `sh` step run there, so `docker compose run` picks up
//   docker-compose.yml + .env (WORKSPACE_DIR_HOST) without an explicit cd.
// * The report container re-mounts the SAME repo from the Docker Desktop VM
//   side (/host_mnt/b/...) - so the SQLite market-data cache and generated
//   HTML persist on the host and are shared with the local uv workflow.
// * No secrets: the pipeline only touches public yfinance data.
// =============================================================================

pipeline {
    // Run on the built-in node (the sidonia Jenkins container, which has the
    // docker CLI + host socket mount). `customWorkspace` pins every stage to
    // the repo root - bind-mounted at /work/chanmainvest/portfolio_dashboard
    // (host B:\) - so `docker compose run` finds docker-compose.yml + .env
    // without an explicit `cd`.
    agent {
        node {
            label 'built-in'
            customWorkspace '/work/chanmainvest/portfolio_dashboard'
        }
    }

    // 03:30 daily. Jenkins cron fields: MINUTE HOUR DOM MONTH DOW.
    triggers { cron('30 3 * * *') }

    options {
        timestamps()                              // prefix every log line with a timestamp
        buildDiscarder(logRotator(numToKeepStr: '30'))  // keep a month of runs
        timeout(time: 30, unit: 'MINUTES')        // cold cache takes minutes, warm seconds
        disableConcurrentBuilds()                 // never two nightly runs at once
    }

    stages {

        // -------------------------------------------------------------------------
        // 0. Build / refresh the report image.
        //
        // `docker compose build report` checks the registry for base-image
        // metadata (python:3.12-slim) on every run. That lookup is the only
        // network dependency outside the report itself, so a transient DNS
        // failure here must not tank the nightly run. We retry a few times
        // and - since the build only needs *some* portfolio-report:latest to
        // exist - treat a registry failure as success when the image is
        // already cached locally. A genuine Dockerfile/dependency error
        // still fails the build.
        // -------------------------------------------------------------------------
        stage('Build report image') {
            steps {
                sh '''set +e
# Fail fast if this build landed in an empty suffixed workspace (@2, @3, ...).
# That happens when a build starts while another still holds the main
# customWorkspace - every compose call would then die with the confusing
# "no configuration file provided: not found" (exit 14).
if [ ! -f docker-compose.yml ]; then
    echo "FATAL: docker-compose.yml not found in $(pwd) - this build got an empty @N workspace while another build holds the real one. Re-run once no portfolio-report-nightly build is active."
    exit 1
fi
for attempt in 1 2 3 4 5; do
    echo "--- build attempt $attempt/5 ---"
    docker compose build report && exit 0
    echo "build attempt $attempt failed (exit $?), waiting 15s before retry..."
    sleep 15
done
# All attempts failed. Proceed only if a portfolio-report:latest already
# exists locally - the report stages reuse it. Otherwise this is a real
# failure.
if docker image inspect portfolio-report:latest >/dev/null 2>&1; then
    echo "WARN: image build failed (likely a transient registry/DNS error), but portfolio-report:latest exists - proceeding with the cached image."
    exit 0
fi
echo "ERROR: image build failed and no portfolio-report:latest is cached. Cannot proceed."
exit 1
'''
            }
        }

        // -------------------------------------------------------------------------
        // 1. Regenerate grandma.html.
        //
        // The container bind-mounts the repo at /app, so it reads
        // grandma_investment_portfolio.xlsx and writes grandma.html +
        // grandma_risk_metrics.json straight into the workspace, sharing the
        // host market_data_cache.db (incremental fetch - a warm nightly run
        // only pulls the newest bars and takes seconds).
        // -------------------------------------------------------------------------
        stage('Regenerate grandma.html') {
            steps {
                sh '''
                    docker compose run --rm report \
                        python build_portfolio_report.py \
                        --input grandma_investment_portfolio.xlsx \
                        --output grandma.html
                    echo "--- generated artifacts ---"
                    ls -l grandma.html grandma_risk_metrics.json
                '''
            }
        }
    }

    // -------------------------------------------------------------------------
    // Always run, regardless of build outcome.
    // -------------------------------------------------------------------------
    post {
        always {
            echo "Pipeline finished: ${currentBuild.currentResult}"
            // Clean up any one-shot report containers the run stages spawned.
            // (Compose `--rm` handles the normal case; this catches strays.)
            sh 'docker compose rm -fsv report 2>/dev/null || true'
        }
        success {
            echo 'grandma.html regenerated.'
        }
        failure {
            echo 'A required stage failed (image build or report generation). See stage logs.'
        }
    }
}
