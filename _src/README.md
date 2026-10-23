원본 묶음(암호화). 복원: `openssl enc -d -aes-256-cbc -pbkdf2 -iter 600000 -in source.tar.gz.enc -pass 'pass:<팀 비밀번호>' | tar -xz`
