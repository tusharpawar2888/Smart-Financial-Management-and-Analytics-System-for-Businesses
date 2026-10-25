pipeline {
    agent any

    environment {
        PYTHON = 'C:\\Users\\HP\\AppData\\Local\\Python\\bin\\python.exe'
    }

    stages {
        stage('Checkout') {
            steps {
                echo 'Source code checked out from GitHub'
            }
        }

        stage('Check Python') {
            steps {
                bat '"%PYTHON%" --version'
                bat '"%PYTHON%" -m pip --version'
            }
        }

        stage('Install Dependencies') {
            steps {
                bat '"%PYTHON%" -m pip install -r requirements.txt'
                bat '"%PYTHON%" -m pip install flake8'
            }
        }

        stage('Code Quality') {
            steps {
                script {
                    def lintStatus = bat(
                        returnStatus: true,
                        script: '"%PYTHON%" -m flake8 app.py --count --statistics'
                    )

                    if (lintStatus != 0) {
                        echo 'WARNING: Flake8 found issues. Review them and improve the code.'
                    } else {
                        echo 'Flake8 passed. No code quality issues found.'
                    }
                }
            }
        }

        stage('Syntax Validation') {
            steps {
                bat '"%PYTHON%" -m py_compile app.py'
                echo 'Python syntax validation passed'
            }
        }

        stage('Complete') {
            steps {
                echo 'Business Finance Desk CI pipeline completed!'
            }
        }
    }
}