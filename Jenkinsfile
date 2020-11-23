def TESTS = [
        '3.7': [
            python: '3.7',
            tox: 'py37'
        ],
        '3.8': [
            python: '3.8',
            tox: 'py38'
        ],
        '3.9': [
            python: '3.9',
            tox: 'py39'
        ]
    ]

pipeline {
    agent none

    stages {
        stage('Code checks') {
            agent {
                docker {
                    image 'python:3.7'
                    args '-u root:sudo'
                }
            }

            steps {
                sh 'pip install tox poetry'
                sh 'tox --parallel--safe-build -e code-checks'
                cleanWs()
            }
        }

        stage('Tests') {
            matrix {
                axes {
                    axis {
                        name 'version'
                        values '3.7', '3.8', '3.9'
                    }
                }

                agent {
                    docker {
                        image "python:${TESTS[version].python}"
                        args '-u root:sudo'
                    }
                }

                stages {
                    stage('Setup') {
                        steps {
                            sh 'pip install tox poetry'
                        }
                    }

                    stage('Tests') {
                        steps {
                            sh "tox --parallel--safe-build -e ${TESTS[version].tox}"
                            cleanWs()
                        }
                    }
                }
            }
        }

        stage('Build doc') {
            agent {
                docker {
                    image 'python:3.7'
                    args '-u root:sudo'
                }
            }

            steps {
                sh 'pip install tox poetry'
                sh 'tox --parallel--safe-build -e doc'
                cleanWs()
            }
        }
    }

    post {
        always {
            sh 'rm -rf .tox'
        }
    }
}