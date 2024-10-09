import random
import string
from docx import Document
import pyperclip

def generate_password(length, use_uppercase, use_lowercase, use_digits, use_symbols,
                      prefix='', must_include='', exclude_chars=''):
    characters = ''
    if use_uppercase:
        characters += string.ascii_uppercase
    if use_lowercase:
        characters += string.ascii_lowercase
    if use_digits:
        characters += string.digits
    if use_symbols:
        characters += string.punctuation

    if must_include:
        if len(must_include) > length:
            raise ValueError("指定された文字列がパスワードの長さを超えています。")
        if any(char in exclude_chars for char in must_include):
            raise ValueError("含める文字列に含まれた文字が除外されている文字セットに含まれています。")

    characters = ''.join(c for c in characters if c not in exclude_chars)
    if not characters and not prefix:
        raise ValueError("使用する文字セットが空です。")

    prefix_length = len(prefix)
    if prefix_length > length:
        raise ValueError("指定された頭文字がパスワードの長さを超えています。")

    remaining_length = length - prefix_length - len(must_include)
    if remaining_length < 0:
        raise ValueError("指定されたパスワードの長さが頭文字と含める文字列の長さを合わせた長さ以下です。")
    
    random_part = ''.join(random.choice(characters) for _ in range(remaining_length))
    password = prefix + random_part
    insert_position = random.randint(0, len(password))
    password = password[:insert_position] + must_include + password[insert_position:]

    return password

def save_password_to_txt(password, file_name):
    if not file_name.endswith('.txt'):
        file_name += '.txt'
    with open(file_name, 'w') as f:
        f.write(password)
    print(f"パスワードが {file_name} に保存されました。")

def save_password_to_docx(password, file_name):
    doc = Document()
    doc.add_paragraph(password)
    if not file_name.endswith('.docx'):
        file_name += '.docx'
    doc.save(file_name)
    print(f"パスワードが {file_name} に保存されました。")

def save_password_to_clipboard(password):
    pyperclip.copy(password)
    print("パスワードがクリップボードに保存されました。")

def main():
    def set_conditions():
        nonlocal length, use_uppercase, use_lowercase, use_digits, use_symbols, prefix, must_include, exclude_chars
        while True:
            try:
                length = int(input("パスワードは何文字にしますか？: "))
                if length <= 0:
                    print("パスワードの長さは1以上の数字を入力してください。")
                else:
                    break
            except ValueError:
                print("有効な数字を入力してください。")

        use_uppercase = input("大文字は使用しますか？ (y/n): ").lower() == 'y'
        use_lowercase = input("小文字は使用しますか？ (y/n): ").lower() == 'y'
        use_digits = input("数字は使用しますか？ (y/n): ").lower() == 'y'
        use_symbols = input("記号は使用しますか？ (y/n): ").lower() == 'y'
        
        option = input("追加のオプションを選んでください。\n1: 頭文字指定\n2: 任意の文字を含む\n3: 任意の文字を含まない\n4: 設定しない\n選択肢: ")
        if option == '1':
            while True:
                prefix = input("頭文字を指定してください: ")
                if len(prefix) > length:
                    print(f"指定された頭文字がパスワードの長さを超えています。残りの文字数は {length - len(prefix)} です。")
                else:
                    must_include = ''
                    exclude_chars = ''
                    break
        elif option == '2':
            while True:
                must_include = input("含めたい文字列を入力してください: ")
                if len(must_include) > length:
                    print("指定した文字列がパスワードの長さを超えています。もう一度入力してください。")
                elif len(must_include) == 0:
                    print("空の文字列は入力できません。")
                else:
                    exclude_chars = ''
                    prefix = ''
                    break
        elif option == '3':
            exclude_chars = input("含めたくない文字列を入力してください: ")
            must_include = ''
            prefix = ''
        elif option == '4':
            prefix = ''
            must_include = ''
            exclude_chars = ''
        else:
            print("1, 2, 3, または 4 を入力してください。")
            prefix = must_include = exclude_chars = ''
    
    length = 0
    use_uppercase = use_lowercase = use_digits = use_symbols = False
    prefix = must_include = exclude_chars = ''

    answer = input("パスワードを作るのかい？ (y/n): ").lower()
    if answer != 'y':
        print("プログラムを終了します。")
        return
    
    set_conditions()
    
    while True:
        try:
            password = generate_password(length, use_uppercase, use_lowercase, use_digits, use_symbols,
                                          prefix, must_include, exclude_chars)
            print("生成されたパスワード:", password)
        except ValueError as e:
            print(e)
            return
        
        while True:
            action = input("項目の数字を入力し操作を選択してください\n1: パスワードを保存する。\n2: 再生成する\n選択肢: ")
            if action == '1':
                while True:
                    save_method = input("項目の数字を入力しパスワードの保存方法を選択してください。\n1: .txtで保存する。\n2: .docxで保存する。\n3: クリップボードに保存する。\n4: クリップボードと.txtまたは.docxに保存する。\n選択肢: ")
                    if save_method in ['1', '2', '3', '4']:
                        break
                    else:
                        print("1, 2, 3, または 4 を入力してください。")

                if save_method in ['1', '2', '4']:
                    file_name = input("ファイルの名前を入力してください（例: mypassword）: ")

                if save_method == '1':
                    save_password_to_txt(password, file_name)
                elif save_method == '2':
                    save_password_to_docx(password, file_name)
                elif save_method == '3':
                    save_password_to_clipboard(password)
                elif save_method == '4':
                    file_format_choice = input("保存したい項目の番号を選んでください。 (1 .txt 2 .docx): ")
                    if file_format_choice == '1':
                        save_password_to_txt(password, file_name)
                    elif file_format_choice == '2':
                        save_password_to_docx(password, file_name)
                    save_password_to_clipboard(password)
                
                return  # 保存後プログラム終了

            elif action == '2':
                set_conditions()  # 条件を再設定する
                break
            else:
                print("1 または 2 を入力してください。")
            
if __name__ == "__main__":
    main()

# python password.py