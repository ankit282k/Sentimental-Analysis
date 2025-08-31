pipeline {
    agent any
    stages {
        stage('Checkout') {
            steps {
                git(
                    url: 'https://github.com/ankit282k/Sentimental-Analysis.git',
                    credentialsId: 'github-cred'
                )
            }
        }
    }
}
