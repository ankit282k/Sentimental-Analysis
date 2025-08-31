pipeline {
    agent any

    environment {
        GIT_CREDENTIALS = 'github-cred'       // Jenkins credential ID for GitHub PAT
        GIT_URL = 'https://github.com/ankit282k/Sentimental-Analysis.git'
        GIT_BRANCH = 'main'
        SCANNER_HOME = tool name: 'SonarQubeScanner', type: 'hudson.plugins.sonar.SonarRunnerInstallation' // ensures correct tool lookup
    }

    stages {

        stage('Checkout SCM') {
            steps {
                checkout([$class: 'GitSCM',
                    branches: [[name: "*/${env.GIT_BRANCH}"]],
                    userRemoteConfigs: [[
                        url: env.GIT_URL,
                        credentialsId: env.GIT_CREDENTIALS
                    ]]
                ])
            }
        }

        stage('Build / Run') {
            steps {
                echo "Running pipeline for branch ${env.GIT_BRANCH}"
                // Run Python environment setup and script
                sh '''
                python3 -m venv venv
                source venv/bin/activate
                pip install -r requirements.txt
                python main.py
                '''
            }
        }

        stage('SonarQube Analysis') {
            steps {
                withSonarQubeEnv('SonarQube') { // Name of Jenkins SonarQube server
                    sh """
                        ${SCANNER_HOME}/bin/sonar-scanner \
                        -Dsonar.projectKey=Sentimental-Analysis \
                        -Dsonar.sources=. \
                        -Dsonar.host.url=${env.SONAR_HOST_URL} \
                        -Dsonar.login=${env.SONAR_AUTH_TOKEN}
                    """
                }
            }
        }

        stage('Quality Gate') {
            steps {
                // Wait for SonarQube analysis to complete
                timeout(time: 1, unit: 'HOURS') {
                    waitForQualityGate abortPipeline: true
                }
            }
        }

    }

    post {
        success {
            echo "Pipeline completed successfully!"
        }
        failure {
            echo "Pipeline failed. Check logs!"
        }
    }
}
