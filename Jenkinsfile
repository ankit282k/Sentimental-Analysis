pipeline {
    agent any

    environment {
        GIT_CREDENTIALS = 'github-cred' // Your Jenkins credential ID
        GIT_URL = 'https://github.com/ankit282k/Sentimental-Analysis.git'
        GIT_BRANCH = 'main'
    }

    stages {

        stage('Checkout SCM') {
            steps {
                // Declarative checkout
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
                echo "Pipeline is running on branch ${env.GIT_BRANCH}"
                // Add your build/test/analysis commands here
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
