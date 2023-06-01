import secrets
import string


def generate(self):
    # define the alphabet
    letters = string.ascii_letters
    digits = string.digits
    special_chars = string.punctuation

    alphabet = letters + digits + special_chars

    # fix password length
    pwd_length = 12

    # generate a password string
    pwd = ''
    for i in range(pwd_length):
        pwd += ''.join(secrets.choice(alphabet))
    print("password string genarated")
    print(pwd)

    # generate password meeting constraints
    while True:
        pwd = ''
        for i in range(pwd_length):
            pwd += ''.join(secrets.choice(alphabet))

        if (any(char in special_chars for char in pwd) and
                sum(char in digits for char in pwd) >= 2):
            break
    print("password meeting constraints:")
    print(pwd)


#checks password strength
def checker(s):
    missing_type = 3
    # if it has at least one lowercase letter, decrease missingTypes by 1
    if any('a' <= c <= 'z' for c in s): missing_type -= 1
    # if it has at least one uppercase letter, decrease missingTypes by 1
    if any('A' <= c <= 'Z' for c in s): missing_type -= 1
    # if it has at least one number, decrease missingTypes by 1
    if any(c.isdigit() for c in s): missing_type -= 1
    change = 0
    one = two = 0
    p = 2
    while p < len(s):
        if s[p] == s[p - 1] == s[p - 2]:
            length = 2
            while p < len(s) and s[p] == s[p - 1]:
                length += 1
                p += 1
            change += length / 3
            if length % 3 == 0:
                one += 1
            elif length % 3 == 1:
                two += 1
        else:
            p += 1
    if len(s) < 7:
        return max(missing_type, 7 - len(s))
    elif len(s) <= 20:
        return max(missing_type, change)
    else:
        delete = len(s) - 20
        change -= min(delete, one)
        change -= min(max(delete - one, 0), two * 2) / 2
        change -= max(delete - one - 2 * two, 0) / 3
        return delete + max(missing_type, change)
