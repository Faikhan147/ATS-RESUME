pipeline {
    agent any

    parameters {
        stashedFile 'RESUME'

        text(
            name: 'JOB_DESCRIPTION',
            defaultValue: '',
            description: 'Paste the complete Job Description'
        )
    }

    environment {
        OPENAI_API_KEY = credentials('openai-api-key')
        AI_MODEL = 'gpt-5.6-luna'
    }

    stages {

        stage('Prepare Workspace') {
            steps {
                sh '''
                    set -e

                    rm -rf output
                    rm -f \
                        resume_original.docx \
                        resume_*.json \
                        ats_*.json \
                        rewrite_*.json \
                        resume_optimized_*.docx \
                        Faisal_Khan_*.docx \
                        jd.txt

                    mkdir -p output

                    printf '%s\\n' "$JOB_DESCRIPTION" > jd.txt
                '''

                unstash 'RESUME'

                sh '''
                    set -e

                    mv RESUME resume_original.docx
                '''
            }
        }


        stage('Validate Resume') {
            steps {
                sh '''
                    set -e

                    echo "======================================"
                    echo "VALIDATING INPUT FILES"
                    echo "======================================"

                    test -f resume_original.docx || {
                        echo "ERROR: Resume DOCX not found"
                        exit 1
                    }

                    test -s jd.txt || {
                        echo "ERROR: Job Description is empty"
                        exit 1
                    }

                    echo ""
                    echo "Resume:"
                    ls -lh resume_original.docx

                    echo ""
                    echo "Resume file type:"
                    file resume_original.docx

                    echo ""
                    echo "Job Description:"
                    wc -c jd.txt

                    echo ""
                    echo "Python:"
                    python3 --version

                    echo ""
                    echo "LibreOffice:"
                    libreoffice --version
                '''
            }
        }


        stage('Install Python Dependencies') {
            steps {
                sh '''
                    set -e

                    python3 -m pip install \
                        --user \
                        --break-system-packages \
                        --upgrade \
                        openai \
                        python-docx
                '''
            }
        }


        stage('Initial Resume Extraction') {
            steps {
                sh '''
                    set -e

                    echo "======================================"
                    echo "EXTRACTING ORIGINAL RESUME"
                    echo "======================================"

                    python3 resume_ai.py extract \
                        resume_original.docx \
                        resume_content.json

                    test -s resume_content.json

                    echo ""
                    echo "Resume extraction completed."
                    ls -lh resume_content.json
                '''
            }
        }


        stage('Initial ATS Analysis') {
            steps {
                sh '''
                    set -e

                    echo "======================================"
                    echo "INITIAL ATS ANALYSIS"
                    echo "======================================"

                    python3 resume_ai.py ats \
                        resume_content.json \
                        jd.txt \
                        ats_initial.json

                    test -s ats_initial.json
                '''

                script {
                    def result = readJSON file: 'ats_initial.json'

                    env.INITIAL_ATS_SCORE = result.ats_score.toString()

                    echo "======================================"
                    echo "INITIAL ATS SCORE = ${env.INITIAL_ATS_SCORE}"
                    echo "======================================"

                    echo "MATCHED SKILLS:"

                    result.matched_skills.each { skill ->
                        echo "  - ${skill}"
                    }

                    echo "MISSING SKILLS:"

                    result.missing_skills.each { skill ->
                        echo "  - ${skill}"
                    }

                    echo "SUGGESTIONS:"

                    result.suggestions.each { suggestion ->
                        echo "  - ${suggestion}"
                    }
                }
            }
        }


        stage('Optimize Resume') {
            steps {
                sh '''
                    set -e

                    echo "======================================"
                    echo "STARTING RESUME OPTIMIZATION"
                    echo "======================================"

                    python3 resume_ai.py optimize \
                        resume_original.docx \
                        jd.txt

                    FINAL_DOCX=$(find . -maxdepth 1 -type f -name 'Faisal_Khan_*.docx' | head -n 1)

                    test -n "$FINAL_DOCX" || {
                        echo "ERROR: Final job-title-based DOCX not found"
                        exit 1
                    }

                    echo ""
                    echo "Best resume generated:"
                    ls -lh "$FINAL_DOCX"
                '''
            }
        }


        stage('Prepare Final Resume') {
            steps {
                sh '''
                    set -e

                    echo "======================================"
                    echo "PREPARING FINAL RESUME"
                    echo "======================================"

                    FINAL_DOCX=$(find . -maxdepth 1 -type f -name 'Faisal_Khan_*.docx' | head -n 1)

                    test -n "$FINAL_DOCX" || {
                        echo "ERROR: Final DOCX not found"
                        exit 1
                    }

                    FINAL_BASENAME=$(basename "$FINAL_DOCX")

                    cp "$FINAL_DOCX" "output/$FINAL_BASENAME"

                    test -f "output/$FINAL_BASENAME"

                    echo ""
                    echo "Final DOCX:"
                    ls -lh "output/$FINAL_BASENAME"
                '''
            }
        }


        stage('Final Resume Extraction') {
            steps {
                sh '''
                    set -e

                    echo "======================================"
                    echo "ANALYZING FINAL RESUME"
                    echo "======================================"

                    python3 resume_ai.py extract \
                        output/resume_final.docx \
                        final_resume_content.json

                    test -s final_resume_content.json
                '''
            }
        }


        stage('Final ATS Analysis') {
            steps {
                sh '''
                    set -e

                    python3 resume_ai.py ats \
                        final_resume_content.json \
                        jd.txt \
                        final_ats_result.json

                    test -s final_ats_result.json
                '''

                script {
                    def result = readJSON file: 'final_ats_result.json'

                    env.FINAL_ATS_SCORE = result.ats_score.toString()

                    echo "======================================"
                    echo "FINAL ATS SCORE = ${env.FINAL_ATS_SCORE}"
                    echo "======================================"

                    echo "FINAL MATCHED SKILLS:"

                    result.matched_skills.each { skill ->
                        echo "  - ${skill}"
                    }

                    echo "FINAL MISSING SKILLS:"

                    result.missing_skills.each { skill ->
                        echo "  - ${skill}"
                    }

                    echo "FINAL SUGGESTIONS:"

                    result.suggestions.each { suggestion ->
                        echo "  - ${suggestion}"
                    }
                }
            }
        }


        stage('Convert DOCX to PDF') {
            steps {
                sh '''
                    set -e

                    echo "======================================"
                    echo "CONVERTING DOCX TO PDF"
                    echo "======================================"

                    FINAL_DOCX=$(find output -maxdepth 1 -type f -name 'Faisal_Khan_*.docx' | head -n 1)

                    test -n "$FINAL_DOCX" || {
                        echo "ERROR: Final DOCX not found in output/"
                        exit 1
                    }

                    libreoffice \
                        --headless \
                        --convert-to pdf \
                        --outdir output \
                        "$FINAL_DOCX"

                    FINAL_PDF="${FINAL_DOCX%.docx}.pdf"

                    test -f "$FINAL_PDF" || {
                        echo "ERROR: PDF conversion failed"
                        exit 1
                    }

                    echo ""
                    echo "Generated files:"
                    ls -lh output/
                '''
            }
        }


        stage('Archive Resume') {
            steps {
                archiveArtifacts artifacts:
                    'output/Faisal_Khan_*.docx,' +
                    'output/Faisal_Khan_*.pdf,' +
                    'final_ats_result.json,' +
                    'ats_initial.json',
                    fingerprint: true
            }
        }
    }


    post {

        success {
            echo "======================================"
            echo "PIPELINE SUCCESS"
            echo "======================================"

            echo "Initial ATS Score: ${env.INITIAL_ATS_SCORE ?: 'N/A'}"
            echo "Final ATS Score:   ${env.FINAL_ATS_SCORE ?: 'N/A'}"

            echo ""
            echo "Final Resume:"
            echo "output/Faisal_Khan_*.docx"

            echo ""
            echo "Final PDF:"
            echo "output/Faisal_Khan_*.pdf"
        }

        failure {
            echo "======================================"
            echo "PIPELINE FAILED"
            echo "======================================"

            echo "Initial ATS Score: ${env.INITIAL_ATS_SCORE ?: 'N/A'}"
            echo "Final ATS Score:   ${env.FINAL_ATS_SCORE ?: 'N/A'}"
        }

        always {
            echo "======================================"
            echo "PIPELINE COMPLETED"
            echo "======================================"
        }
    }
}
