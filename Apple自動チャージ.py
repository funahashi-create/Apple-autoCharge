import re, time, random, datetime, json, os, ctypes, sys
import tkinter as tk
from tkinter import messagebox
import psutil
import mss
from PIL import Image, ImageChops
import gspread
from pywinauto import Desktop
from pywinauto.keyboard import send_keys
from pywinauto.uia_defines import IUIA
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.uia_element_info import UIAElementInfo

# ===== 設定 =====
CREDENTIALS_JSON = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "service_account.json")
SPREADSHEET_KEY = "17Vb12hNhiJ_ehilOyjMutUemsDnbcqvrY9C6PUJ5k9Q"
SHEET_NAMES = ["満額コードチャージ", "残額コードチャージ"]
SHEET_OFFSET = {"満額コードチャージ": 0, "残額コードチャージ": 1}  # B列に管理番号がある分ずらす
TARGET_EXE = "applemusic.exe"
RESULT_TIMEOUT = 30   # チャージ結果を待つ最大秒数
NEXT_TIMEOUT = 10     # 「他のコードを使用」後に入力欄が戻るまでの最大秒数
REGION_FILE = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "redeem_region.json")
RED_TH = 0.01         # 枠内の赤ピクセル割合（エラー判定）
CHANGE_TH = 0.25      # 基準画像からの変化割合（画面切替判定）
UIA_FALLBACK = 1.0    # 画像で検知できなくてもUIA確認する間隔(秒)
RETRY_FIRST = 10      # 弾かれた後の1回目の再送までの秒数
RETRY_MIN = 8         # 2回目以降の再送間隔 最小(秒)
RETRY_MAX = 15        # 2回目以降の再送間隔 最大(秒)
RETRY_LIMIT = 120     # 弾かれてからこの秒数を超えたらエラー停止
REJECT_GRACE = 3.0    # 押下後この秒数、結果なし＆ボタン有効のままなら弾かれ扱い
SEND_TIMES = []       # 受理された送信時刻（ログ用）
SIGNIN_TIMEOUT = 20   # 再ログイン後、入力画面に戻るまでの最大秒数

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    pass


def popup(msg, error=False):
    r = tk.Tk(); r.withdraw(); r.attributes("-topmost", True)
    (messagebox.showerror if error else messagebox.showinfo)("自動チャージ", msg, parent=r)
    r.destroy()


def ask_continue(msg, win=None):
    root = tk.Tk(); root.title("自動チャージ エラー（一時停止中）"); root.attributes("-topmost", True)
    body = tk.Frame(root); body.pack(padx=16, pady=12)
    icon = tk.Canvas(body, width=36, height=36, highlightthickness=0, bg=body.cget("bg"))
    icon.create_oval(2, 2, 34, 34, fill="#D0453A", outline="")
    icon.create_text(18, 18, text="!", fill="white", font=("", 18, "bold"))
    icon.pack(side="left", anchor="n", padx=(0, 12))
    tk.Label(body, text=msg, justify="left").pack(side="left")
    choice = {"go": False}

    def pick(v):
        choice["go"] = v; root.destroy()

    f = tk.Frame(root); f.pack(pady=(0, 12))
    tk.Button(f, text="次のコードから継続", width=18, command=lambda: pick(True)).pack(side="left", padx=6)
    tk.Button(f, text="チャージ終了", width=12, command=lambda: pick(False)).pack(side="left", padx=6)
    root.protocol("WM_DELETE_WINDOW", lambda: pick(False))

    if win is not None:  # Apple Musicウィンドウの左隣（余裕がなければ左上に少し重ねる）
        try:
            r = win.rectangle()
            root.update_idletasks()
            w = root.winfo_reqwidth() + 20  # 枠線分の余白
            screen_left = ctypes.windll.user32.GetSystemMetrics(76)  # 仮想画面の左端（マルチモニター対応）
            if r.left - w >= screen_left:
                x, y = r.left - w, r.top
            else:
                x, y = r.left + 20, r.top + 60
            root.geometry(f"+{x}+{max(y, 0)}")
        except Exception:
            pass
    root.mainloop()
    return choice["go"]


def get_account(win):
    try:
        el = win.child_window(title_re=".*に適用されます.*", control_type="Text", found_index=0)
        if el.exists(timeout=3):
            m = re.search(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+", el.window_text(), re.ASCII)
            if m:
                return m.group(0)
    except Exception:
        pass
    return None


def ask_account():
    root = tk.Tk(); root.title("適用アカウント"); root.attributes("-topmost", True)
    tk.Label(root, text="適用アカウントを画面から取得できませんでした。\n"
                        "メールアドレスを入力するか、指定なしで開始してください。",
             justify="left", padx=16, pady=12).pack()
    entry = tk.Entry(root, width=36); entry.pack(padx=16)
    result = {"v": None}

    def by_input():
        v = entry.get().strip()
        if "@" not in v:
            messagebox.showerror("入力エラー", "メールアドレスを入力してください", parent=root); return
        result["v"] = v; root.destroy()

    def no_account():
        result["v"] = "指定なし"; root.destroy()

    f = tk.Frame(root); f.pack(pady=12)
    tk.Button(f, text="入力したアドレスで開始", command=by_input).pack(side="left", padx=6)
    tk.Button(f, text="アカウント指定なしで開始", command=no_account).pack(side="left", padx=6)
    tk.Button(f, text="キャンセル", width=10, command=root.destroy).pack(side="left", padx=6)
    entry.focus_force()
    root.mainloop()
    return result["v"]


def recover(win):
    try:
        ok_btn = win.child_window(auto_id="PrimaryButton", control_type="Button", found_index=0)
        if ok_btn.exists(timeout=2):
            ok_btn.invoke()
            ok_btn.wait_not("exists", timeout=5)
        else:
            nxt = win.child_window(auto_id="redeemAnotherButton", control_type="Button", found_index=0)
            if nxt.exists(timeout=2):
                nxt.invoke()
        return win.child_window(auto_id="codeField", control_type="Edit", found_index=0).exists(timeout=NEXT_TIMEOUT)
    except Exception:
        return False


def load_region():
    try:
        with open(REGION_FILE, encoding="utf-8") as f:
            rel = json.load(f)
        return rel if isinstance(rel, list) and len(rel) == 4 else None
    except Exception:
        return None


def save_region(rel):
    with open(REGION_FILE, "w", encoding="utf-8") as f:
        json.dump(rel, f)


def select_region(root, win):
    try:
        win.set_focus()
    except Exception:
        pass
    time.sleep(0.8)
    r = win.rectangle()
    top = tk.Toplevel(root)
    top.attributes("-fullscreen", True); top.attributes("-alpha", 0.3); top.attributes("-topmost", True)
    cv = tk.Canvas(top, bg="black", highlightthickness=0, cursor="crosshair")
    cv.pack(fill="both", expand=True)
    cv.create_text(20, 20, anchor="nw", fill="white", font=("", 16),
                   text="検知範囲をドラッグで指定（Escでキャンセル）")
    st = {}

    def down(e):
        st["sx"], st["sy"] = e.x, e.y
        st["rect"] = cv.create_rectangle(e.x, e.y, e.x, e.y, outline="yellow", width=3)

    def move(e):
        if "rect" in st:
            cv.coords(st["rect"], st["sx"], st["sy"], e.x, e.y)

    def up(e):
        if "sx" not in st:
            return
        x1, x2 = sorted((st["sx"], e.x)); y1, y2 = sorted((st["sy"], e.y))
        if x2 - x1 > 20 and y2 - y1 > 20:
            ox, oy = top.winfo_rootx(), top.winfo_rooty()
            st["rel"] = [ox + x1 - r.left, oy + y1 - r.top, ox + x2 - r.left, oy + y2 - r.top]
        top.destroy()

    cv.bind("<ButtonPress-1>", down); cv.bind("<B1-Motion>", move); cv.bind("<ButtonRelease-1>", up)
    top.bind("<Escape>", lambda e: top.destroy())
    top.focus_force(); top.grab_set()
    root.wait_window(top)
    return st.get("rel")


def ask_settings(win):
    root = tk.Tk(); root.title("Appleギフトコード自動チャージ"); root.attributes("-topmost", True)
    region = {"rel": load_region()}
    sheet = tk.StringVar(value=SHEET_NAMES[0])
    for i, name in enumerate(SHEET_NAMES):
        tk.Radiobutton(root, text=name, variable=sheet, value=name).grid(row=0, column=i, sticky="w", padx=8, pady=6)
    fields = [("batch", "1バッチの件数", "40"), ("rest", "バッチ後の休憩(秒)", "60"),
              ("wmin", "完了後待機 最小(秒)", "1.5"), ("wmax", "完了後待機 最大(秒)", "2.0"),
              ("total", "総件数上限(0=全件)", "0")]
    entries = {}
    for r, (key, label, default) in enumerate(fields, start=1):
        tk.Label(root, text=label).grid(row=r, column=0, sticky="w", padx=8)
        e = tk.Entry(root, width=10); e.insert(0, default); e.grid(row=r, column=1, padx=8, pady=2)
        entries[key] = e
    pw_row = len(fields) + 1
    tk.Label(root, text="パスワード(再ログイン用)").grid(row=pw_row, column=0, sticky="w", padx=8)
    pw_entry = tk.Entry(root, width=20, show="*"); pw_entry.grid(row=pw_row, column=1, padx=8, pady=2)
    result = {}

    def ok():
        try:
            cfg = {"sheet": sheet.get(),
                   "batch": int(entries["batch"].get()), "rest": float(entries["rest"].get()),
                   "wmin": float(entries["wmin"].get()), "wmax": float(entries["wmax"].get()),
                   "total": int(entries["total"].get())}
            if cfg["batch"] < 1 or cfg["wmin"] < 0 or cfg["wmin"] > cfg["wmax"] or cfg["total"] < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("入力エラー", "数値を確認してください", parent=root); return
        if not region["rel"]:
            messagebox.showerror("範囲未設定", "「範囲を選択」で検知範囲を指定してください", parent=root); return
        cfg["region"] = region["rel"]
        cfg["password"] = pw_entry.get()
        result.update(cfg); root.destroy()

    n = len(fields) + 2
    region_label = tk.Label(root, text="検知範囲: " + ("設定済み" if region["rel"] else "未設定"))
    region_label.grid(row=n, column=0, sticky="w", padx=8)

    def pick_region():
        root.withdraw()
        rel = select_region(root, win)
        root.deiconify()
        if rel:
            region["rel"] = rel
            save_region(rel)
            region_label.config(text="検知範囲: 設定済み")

    tk.Button(root, text="範囲を選択", width=10, command=pick_region).grid(row=n, column=1, pady=4)
    n += 1
    tk.Button(root, text="開始", width=10, command=ok).grid(row=n, column=0, pady=8)
    tk.Button(root, text="キャンセル", width=10, command=root.destroy).grid(row=n, column=1, pady=8)
    root.mainloop()
    return result or None


def get_redeem_window():
    for w in Desktop(backend="uia").windows():
        try:
            if psutil.Process(w.element_info.process_id).name().lower() != TARGET_EXE:
                continue
        except Exception:
            continue
        spec = Desktop(backend="uia").window(handle=w.handle)
        if spec.child_window(auto_id="codeField", control_type="Edit", found_index=0).exists(timeout=1):
            return spec
    return None


def read_value(field):
    try:
        return (field.get_value() or "").strip()
    except Exception:
        return None


def wait_value(field, cond, timeout=1.5):
    end = time.time() + timeout
    while time.time() < end:
        v = read_value(field)
        if v is None or cond(v):
            return True
        time.sleep(0.1)
    return False


def clear_field(field):
    for _ in range(3):
        field.click_input()
        time.sleep(0.8)
        send_keys("{END}{BACKSPACE 40}", pause=0.01)
        if wait_value(field, lambda v: v == ""):
            return True
    return False


def make_box(win, rel):
    r = win.rectangle()
    return {"left": r.left + rel[0], "top": r.top + rel[1],
            "width": rel[2] - rel[0], "height": rel[3] - rel[1]}


def grab(sct, box):
    s = sct.grab(box)
    return Image.frombytes("RGB", s.size, s.rgb).resize((96, 64))


def red_ratio(img):
    px = img.getdata()
    return sum(1 for r, g, b in px if r > 180 and g < 90 and b < 110) / len(px)


def change_ratio(base, img):
    hist = ImageChops.difference(base, img).convert("L").histogram()
    return sum(hist[41:]) / (img.width * img.height)


_UIA = IUIA()
_RESULT_COND = _UIA.iuia.CreateOrCondition(
    _UIA.iuia.CreatePropertyCondition(_UIA.UIA_dll.UIA_AutomationIdPropertyId, "AskUserExplanation"),
    _UIA.iuia.CreatePropertyCondition(_UIA.UIA_dll.UIA_AutomationIdPropertyId, "resultHeaderText"),
)
_REDEEM_COND = _UIA.iuia.CreatePropertyCondition(_UIA.UIA_dll.UIA_AutomationIdPropertyId, "redeemButton")
_TEXT_COND = _UIA.iuia.CreatePropertyCondition(_UIA.UIA_dll.UIA_ControlTypePropertyId,
                                               _UIA.UIA_dll.UIA_TextControlTypeId)


def find_id(win, auto_id, timeout=0):
    """AutomationIdの要素をネイティブ検索で探す（0.1秒間隔、見つからなければNone）"""
    cond = _UIA.iuia.CreatePropertyCondition(_UIA.UIA_dll.UIA_AutomationIdPropertyId, auto_id)
    end = time.time() + timeout
    while True:
        try:
            el = win.wrapper_object().element_info.element.FindFirst(_UIA.tree_scope["descendants"], cond)
            if el:
                return UIAWrapper(UIAElementInfo(el))
        except Exception:  # フォールバック（従来方式）
            spec = win.child_window(auto_id=auto_id, found_index=0)
            return spec.wrapper_object() if spec.exists(timeout=max(0, end - time.time())) else None
        if time.time() >= end:
            return None
        time.sleep(0.1)


def find_charge_text(hdr, timeout=3):
    """「コードが適用されました」の近く（親・その親）から「チャージしました」の文言を探す"""
    end = time.time() + timeout
    while True:
        try:
            node = hdr
            for _ in range(2):
                node = _UIA.iuia.ControlViewWalker.GetParentElement(node)
                arr = node.FindAll(_UIA.tree_scope["descendants"], _TEXT_COND)
                for i in range(arr.Length):
                    name = arr.GetElement(i).CurrentName or ""
                    if "チャージしました" in name:
                        return name
        except Exception:
            return None
        if time.time() >= end:
            return None
        time.sleep(0.1)


def check_result(win, root=None):
    hit, hdr = None, None
    if root is not None:  # ネイティブ検索（エラー・成功を1回で探す）
        try:
            el = root.FindFirst(_UIA.tree_scope["descendants"], _RESULT_COND)
            hit = (el.CurrentAutomationId, el.CurrentName) if el else ("", "")
            hdr = el if el else None
        except Exception:
            hit = None
    if hit is None:  # フォールバック（従来方式）
        err = win.child_window(auto_id="AskUserExplanation", found_index=0)
        if err.exists(timeout=0):
            hit = ("AskUserExplanation", err.window_text())
        elif win.child_window(auto_id="resultHeaderText", found_index=0).exists(timeout=0):
            hit = ("resultHeaderText", "")
        else:
            hit = ("", "")
    aid, name = hit
    if aid == "AskUserExplanation":
        return "error", name, None
    if aid == "resultHeaderText":
        msg = find_charge_text(hdr) if hdr is not None else None
        if msg is None:  # フォールバック（従来方式）
            el = win.child_window(title_re=".*チャージしました.*", control_type="Text", found_index=0)
            msg = el.window_text() if el.exists(timeout=3) else "コードが適用されました。"
        m = re.search(r"¥([\d,]+)をチャージ", msg)
        return "ok", msg, int(m.group(1).replace(",", "")) if m else None
    return None


def btn_enabled(btn):
    try:
        return btn.is_enabled()
    except Exception:
        return None  # ボタンが消えた・読めない


def recent_sends(sec=60):
    now = time.time()
    return sum(1 for t in SEND_TIMES if now - t <= sec)


def window_wait(sec=60, limit=5, margin=2.0):
    """直近sec秒の受理がlimit件以上なら、枠が1件空く時刻まで待つ秒数を返す"""
    now = time.time()
    recent = sorted(t for t in SEND_TIMES if now - t <= sec)
    if len(recent) < limit:
        return RETRY_FIRST  # 想定ルールと合わない場合は従来どおり
    return max(1.0, recent[len(recent) - limit] + sec + margin - now)


def type_slow(text):
    """1文字ずつ0.1〜0.25秒間隔で入力（send_keysの特殊文字はエスケープ）"""
    for ch in text:
        send_keys("{%s}" % ch if ch in "+^%~(){}[]" else ch, pause=0, with_spaces=True)
        time.sleep(random.uniform(0.1, 0.25))


def sign_in(win, account, password):
    """サインイン画面があれば再ログイン。None=画面なし / (True, msg)=成功 / (False, msg)=失敗"""
    pw = find_id(win, "passwordTextBox")
    if pw is None:
        return None
    if not password:
        return False, "サインイン画面が表示されました（パスワード未設定・未適用）"
    win.set_focus()
    user = find_id(win, "userNameTextBox")
    uv = read_value(user) if user is not None else None
    named = account and account != "指定なし"
    if uv == "":
        if not named:
            return False, "サインイン画面のアドレスが空です（未適用）"
        user.click_input(); time.sleep(0.5)
        type_slow(account)
    elif uv and named and uv.lower() != account.lower():
        return False, f"サインイン画面のアカウントが異なります（未適用）: {uv}"
    pw.click_input(); time.sleep(0.5)
    send_keys("^a{BACKSPACE}", pause=0.05)
    type_slow(password)
    btn = find_id(win, "continueButton", timeout=3)
    end = time.time() + 5
    while btn is not None and btn_enabled(btn) is not True and time.time() < end:
        time.sleep(0.1)
    if btn is None or btn_enabled(btn) is not True:
        return False, "サインインの「続ける」が有効になりません（未適用）"
    btn.invoke()
    end = time.time() + SIGNIN_TIMEOUT
    while time.time() < end:
        if find_id(win, "passwordTextBox") is None and find_id(win, "codeField") is not None:
            return True, "成功"
        time.sleep(0.3)
    return False, "サインイン後に入力画面へ戻れませんでした（未適用）"


def redeem_one(win, code, sct, rel, label="", account="", password=""):
    field = find_id(win, "codeField", timeout=5)
    if field is None:
        return "error", "入力画面が見つかりません", None
    win.set_focus()
    if not clear_field(field):
        return "error", "入力欄をクリアできませんでした（未送信）", None
    send_keys(code, pause=0.02)
    if not wait_value(field, lambda v: v.upper() == code.upper()):
        return "error", f"入力値がシートと一致しません（未送信）: {read_value(field)}", None
    box = make_box(win, rel)
    base = grab(sct, box)
    try:
        root = win.wrapper_object().element_info.element
    except Exception:
        root = None
    btn = None
    if root is not None:
        try:
            el = root.FindFirst(_UIA.tree_scope["descendants"], _REDEEM_COND)
            btn = UIAWrapper(UIAElementInfo(el)) if el else None
        except Exception:
            btn = None
    if btn is None:  # フォールバック（従来方式）
        btn = win.child_window(auto_id="redeemButton", control_type="Button", found_index=0)
    ready_end = time.time() + 3
    while True:
        try:
            if btn.is_enabled():
                break
        except Exception:
            break  # 状態が読めない場合は従来どおり押す
        if time.time() >= ready_end:
            return "error", "「コードを使う」が有効になりませんでした（未送信）", None
        time.sleep(0.1)
    btn.invoke()
    sent = time.time()
    rejected_at, tries, last_wait, was_disabled = None, 0, 0.0, False

    def finish(res):
        SEND_TIMES.append(sent)
        if rejected_at is not None:
            print(f"[通過] {tries}回目 / 弾かれてから{time.time() - rejected_at:.1f}秒")
        return res

    end = sent + RESULT_TIMEOUT
    next_uia = sent + UIA_FALLBACK
    while time.time() < end:
        img = grab(sct, box)
        if red_ratio(img) > RED_TH or change_ratio(base, img) > CHANGE_TH or time.time() >= next_uia:
            res = check_result(win, root)
            if res:
                return finish(res)
            next_uia = time.time() + UIA_FALLBACK
        en = btn_enabled(btn)
        if en is False:
            was_disabled = True
        elif en and (was_disabled or time.time() - sent >= REJECT_GRACE):
            time.sleep(0.3)  # 結果画面への切替途中を誤検知しないよう再確認
            res = check_result(win, root)
            if res:
                return finish(res)
            si = sign_in(win, account, password)
            if si is not None:
                si_ok, smsg = si
                print(f"[サインイン] {datetime.datetime.now():%H:%M:%S} 再ログイン → {smsg}")
                if not si_ok:
                    return "error", smsg, None
                field = find_id(win, "codeField", timeout=5)
                btn = find_id(win, "redeemButton", timeout=5)
                if field is None or btn is None or not wait_value(field, lambda x: x.upper() == code.upper()):
                    return "error", "サインイン後の入力画面を確認できません（未送信）", None
                time.sleep(1.0)
                base = grab(sct, box)
                btn.invoke()
                sent, was_disabled = time.time(), False
                end, next_uia = sent + RESULT_TIMEOUT, sent + UIA_FALLBACK
                time.sleep(0.05)
                continue
            v = read_value(field)
            if v is not None and v.upper() == code.upper():
                if rejected_at is None:
                    rejected_at = time.time()
                    print(f"[弾かれ] {datetime.datetime.now():%H:%M:%S} {label} (直前60秒の送信: {recent_sends()}件)")
                else:
                    print(f"[再送] {tries}回目 {last_wait:.1f}秒後 → 弾かれ")
                tries += 1
                last_wait = window_wait() if tries == 1 else random.uniform(RETRY_MIN, RETRY_MAX)
                if time.time() - rejected_at + last_wait > RETRY_LIMIT:
                    return "error", f"{RETRY_LIMIT}秒再送しても弾かれたため停止（未適用）", None
                time.sleep(last_wait)
                if not wait_value(field, lambda x: x.upper() == code.upper()):
                    return "error", f"再送前の入力値がシートと一致しません（未送信）: {read_value(field)}", None
                base = grab(sct, box)
                btn.invoke()
                sent, was_disabled = time.time(), False
                end, next_uia = sent + RESULT_TIMEOUT, sent + UIA_FALLBACK
        time.sleep(0.05)
    return "timeout", "応答タイムアウト（チャージ済みの可能性あり・要確認）", None


def go_next(win):
    btn = find_id(win, "redeemAnotherButton", timeout=5)
    if btn is None:
        return False
    btn.invoke()
    return find_id(win, "codeField", timeout=NEXT_TIMEOUT) is not None


def to_int(text):
    digits = re.sub(r"[^\d]", "", text or "")
    return int(digits) if digits else None


def main():
    win = get_redeem_window()
    if not win:
        popup("Apple Musicの「コードを使う」画面を開いてから実行してください。", True); return
    cfg = ask_settings(win)
    if not cfg:
        return
    account = get_account(win) or ask_account()
    if not account:
        return
    print(f"適用アカウント: {account}")
    sct = mss.mss()
    ws = gspread.service_account(filename=CREDENTIALS_JSON).open_by_key(SPREADSHEET_KEY).worksheet(cfg["sheet"])
    rows = ws.get_all_values()
    off = SHEET_OFFSET.get(cfg["sheet"], 0)
    face_i, st_i = 1 + off, 2 + off  # 額面・ステータスの列番号（0始まり）
    st_col, end_col, acc_col = (chr(ord("A") + st_i + k) for k in (0, 3, 4))

    if any(len(r) > st_i and r[st_i].strip() == "処理中" for r in rows[1:]):
        popup("「処理中」の行があります。\nApple側の残高を確認して手動で修正してから再実行してください。", True)
        return

    targets = [(i, r) for i, r in enumerate(rows[1:], start=2)
               if r and r[0].strip() and (len(r) <= st_i or not r[st_i].strip())]
    if cfg["total"] > 0:
        targets = targets[:cfg["total"]]
    if not targets:
        popup("未処理のコードがありません。"); return

    def write(row, status, msg, amount):
        now = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        ws.update(range_name=f"{st_col}{row}:{end_col}{row}", values=[[status, msg, "" if amount is None else amount, now]])

    done, total_amount, errors = 0, 0, 0

    def pause(row, msg):
        nonlocal errors
        errors += 1
        go = ask_continue(f"{row}行目でエラーのため一時停止しました。\n{msg}\n\n"
                          f"完了: {done}件 / 合計 ¥{total_amount:,}", win)
        if go and not recover(win):
            popup("画面を入力状態に戻せなかったため終了します。", True)
            return False
        return go

    for n, (row, r) in enumerate(targets, start=1):
        code = r[0].strip()
        face = to_int(r[face_i]) if len(r) > face_i else None
        ws.batch_update([{"range": f"{st_col}{row}", "values": [["処理中"]]},
                         {"range": f"{acc_col}{row}", "values": [[account]]}])
        try:
            status, msg, amount = redeem_one(win, code, sct, cfg["region"], f"{row}行目", account, cfg["password"])
        except Exception as e:
            status, msg, amount = "error", f"想定外の例外（要確認）: {e}", None

        if status == "ok" and face is not None and amount != face:
            status, msg = "mismatch", f"額面不一致（チャージは実行済み）: {msg}"
            total_amount += amount or 0

        if status != "ok":
            print(f"[{n}/{len(targets)}] {row}行目 エラー: {msg}")
            go = pause(row, msg)
            write(row, "エラー", msg, amount)
            if not go:
                break
            continue

        write(row, "完了", msg, amount)
        done += 1
        total_amount += amount or 0
        print(f"[{n}/{len(targets)}] {row}行目 完了 ¥{amount}")

        if n == len(targets):
            break
        if not go_next(win):
            if not pause(row, "完了後、次の入力画面に戻れませんでした（この行のチャージは完了済み）"):
                break
        time.sleep(random.uniform(cfg["wmin"], cfg["wmax"]))
        if n % cfg["batch"] == 0:
            print(f"--- {cfg['batch']}件完了、{cfg['rest']}秒休憩 ---")
            time.sleep(cfg["rest"])

    popup(f"終了しました。\n完了: {done}件 / エラー: {errors}件 / 合計 ¥{total_amount:,}")


if __name__ == "__main__":
    main()