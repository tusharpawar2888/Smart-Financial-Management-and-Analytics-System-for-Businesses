pipeline {
    agent any

    stages {
        stage('Checkout') {
            steps {
                echo 'Source code checked out from GitHub'
            }
        }

        stage('Check Python') {
            steps {
                bat 'python --version'
                bat 'python -m pip --version'
            }
        }

        stage('Install Dependencies') {
            steps {
                bat 'python -m pip install -r requirements.txt'
                bat 'python -m pip install flake8'
            }
        }

        stage('Code Quality') {
            steps {
                bat 'python -m flake8 app.py --count --statistics'
            }
        }

        stage('Syntax Validation') {
            steps {
                bat 'python -m py_compile app.py'
            }
        }

        stage('Complete') {
            steps {
                echo 'Business Finance Desk CI pipeline completed!'
            }
        }
    }
}