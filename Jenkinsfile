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
                    rm -rf output
                    mkdir -p output

                    printf '%s\\n' "$JOB_DESCRIPTION" > jd.txt
                '''

                unstash 'RESUME'

                sh '''
                    mv RESUME resume_original.docx
                '''
            }
        }

        stage('Validate Resume') {
            steps {
                sh '''
                    test -f resume_original.docx || {
                        echo "ERROR: Resume DOCX not found"
                        exit 1
                    }

                    test -s jd.txt || {
                        echo "ERROR: Job Description is empty"
                        exit 1
                    }

                    file resume_original.docx

                    echo "Job Description saved successfully:"
                    wc -c jd.txt

                    python3 --version
                    libreoffice --version
                '''
            }
        }

        stage('Install Python Dependencies') {
            steps {
                sh '''
                    python3 -m pip install --user --upgrade openai python-docx
                '''
            }
        }

        stage('Extract Resume Content') {
            steps {
                sh '''
                    python3 resume_ai.py extract resume_original.docx resume_content.json
                '''
            }
        }

        stage('ATS Analysis') {
            steps {
                sh '''
                    python3 resume_ai.py ats \
                        resume_content.json \
                        jd.txt \
                        ats_result.json
                '''

                script {
                    def result = readJSON file: 'ats_result.json'

                    env.ATS_SCORE = result.ats_score.toString()

                    echo "======================================"
                    echo "ATS SCORE = ${env.ATS_SCORE}"
                    echo "======================================"

                    echo "MATCHED SKILLS:"
                    echo result.matched_skills.join(', ')

                    echo "MISSING SKILLS:"
                    echo result.missing_skills.join(', ')

                    echo "SUGGESTIONS:"
                    echo result.suggestions.join('\n')
                }
            }
        }

        stage('Decide Resume Action') {
            steps {
                script {

                    if (env.ATS_SCORE.toInteger() < 80) {

                        echo "ATS score is below 80."
                        echo "Resume rewriting will start."

                    } else {

                        echo "ATS score is 80 or above."
                        echo "Original resume will be kept."
                    }
                }
            }
        }

        stage('Rewrite Resume') {

            when {
                expression {
                    env.ATS_SCORE.toInteger() < 80
                }
            }

            steps {
                sh '''
                    python3 resume_ai.py rewrite \
                        resume_original.docx \
                        resume_content.json \
                        jd.txt \
                        rewritten_content.json

                    python3 resume_ai.py apply \
                        resume_original.docx \
                        rewritten_content.json \
                        output/resume_v2.docx
                '''
            }
        }

        stage('Keep Original Resume') {

            when {
                expression {
                    env.ATS_SCORE.toInteger() >= 80
                }
            }

            steps {
                sh '''
                    cp resume_original.docx output/resume_v2.docx
                '''
            }
        }

        stage('Convert DOCX to PDF') {
            steps {
                sh '''
                    libreoffice \
                        --headless \
                        --convert-to pdf \
                        --outdir output \
                        output/resume_v2.docx

                    ls -lh output/
                '''
            }
        }

        stage('Archive Resume') {
            steps {
                archiveArtifacts artifacts: 'output/resume_v2.docx,output/resume_v2.pdf,ats_result.json',
                    fingerprint: true
            }
        }
    }

    post {
        always {
            echo "Pipeline completed."
            echo "ATS Score: ${env.ATS_SCORE ?: 'N/A'}"
        }
    }
}
