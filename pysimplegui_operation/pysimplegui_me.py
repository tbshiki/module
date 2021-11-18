
def check_data(name, password) -> bool:
    """
    ログイン名とパスワードが、指定された値と合致しているか判定する関数

    Args:
        name (str): 入力された名前
        password (str): 入力されたパスワード

    Returns:
        bool: 判定結果
    """
    # 正しいと判定するデータ
    correct_data = ("TERU", "TERU_PASSWORD")

    if (name, password) == correct_data:
        return True
    return False