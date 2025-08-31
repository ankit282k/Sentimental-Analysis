pipeline {
    agent {
        docker {
            image 'python:3.10'
        }
    }

    stages {
        stage('Checkout') {
            steps {
                checkout scm
            }
        }

        stage('Setup Python') {
            steps {
                sh 'python3 --version'
                sh 'pip install --upgrade pip'
                // If you have dependencies, uncomment below:
                // sh 'pip install -r requirements.txt'
            }
        }

        stage('SonarQube Analysis') {
            steps {
                withSonarQubeEnv('MySonarQubeServer') {
                    sh 'sonar-scanner'
                }
            }
        }

        stage('Quality Gate') {
            steps {
                timeout(time: 1, unit: 'MINUTES') {
                    waitForQualityGate abortPipeline: true
                }
            }
        }
    }
}
