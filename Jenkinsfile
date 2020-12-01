
def test_environments = ['py37', 'py38', 'py39']

pipeline {
    agent any

    stages {
        stage("Code checks") {
            steps {
                sh 'tox -e code-checks'
            }
        }

        stage("Tests") {
            steps {
                script {
                    def commands = test_environments.collectEntries({ tox_env ->
                        [tox_env, {
                            sh "tox --workdir .tox_$tox_env -vv -e $tox_env"
                            }
                        ]
                    })
                    parallel(commands)
                }
            }
        }
    }
}