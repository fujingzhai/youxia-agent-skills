# -*- coding: utf-8 -*-
"""知网 AJ-Captcha 拼图滑块：边缘模板匹配定位缺口 + 先加速后减速的仿真拖动"""
import base64, random
import numpy as np, cv2


def _b64img(src):
    return cv2.imdecode(np.frombuffer(base64.b64decode(src.split(",", 1)[1]), np.uint8),
                        cv2.IMREAD_UNCHANGED)


def present(page):
    return page.query_selector(".verify-move-block") is not None


def solve(page, tries=6, verbose=True):
    """页面上若有滑块则破之。返回 True 表示当前页面已无滑块。"""
    for i in range(tries):
        if not present(page):
            return True
        try:    # 滑块框先于拼图出现，图没到齐就匹配只会瞎拖
            page.wait_for_function("""() => ['.verify-img-panel img', '.verify-sub-block img'].every(
                s => (document.querySelector(s)?.getAttribute('src') || '').startsWith('data:image'))""",
                timeout=8000)
            page.wait_for_timeout(300)
            bg = _b64img(page.get_attribute(".verify-img-panel img", "src"))
            tp = _b64img(page.get_attribute(".verify-sub-block img", "src"))
        except Exception:
            return False
        if tp.shape[2] == 4:                      # 裁掉滑块图的透明边
            ys, xs = np.where(tp[:, :, 3] > 10)
            tp = tp[ys.min():ys.max() + 1, xs.min():xs.max() + 1]
        e1 = cv2.Canny(cv2.cvtColor(bg[:, :, :3], cv2.COLOR_BGR2GRAY), 100, 200)
        e2 = cv2.Canny(cv2.cvtColor(tp[:, :, :3], cv2.COLOR_BGR2GRAY), 100, 200)
        _, score, _, loc = cv2.minMaxLoc(cv2.matchTemplate(e1, e2, cv2.TM_CCOEFF_NORMED))
        panel = page.query_selector(".verify-img-panel").bounding_box()
        target = loc[0] * panel["width"] / bg.shape[1]

        blk = page.query_selector(".verify-move-block").bounding_box()
        sx, sy = blk["x"] + blk["width"] / 2, blk["y"] + blk["height"] / 2
        page.mouse.move(sx, sy); page.mouse.down()
        cur = v = 0.0; mid = target * 0.75
        while cur < target:
            acc = 2.5 if cur < mid else -3.0
            v0 = v; v = max(0.5, v + acc * 0.3)
            cur += max(1.0, v0 * 0.3 + 0.5 * acc * 0.09)
            page.mouse.move(sx + min(cur, target), sy + random.uniform(-1.5, 1.5))
            page.wait_for_timeout(random.randint(8, 22))
        page.wait_for_timeout(random.randint(150, 350))
        page.mouse.move(sx + target, sy); page.mouse.up()
        page.wait_for_timeout(2500)

        ok = not present(page)
        if verbose:
            print(f"  滑块       第{i+1}次 匹配度{score:.2f} → {'通过' if ok else '未过，换图重试'}")
        if ok:
            return True
        r = page.query_selector(".verify-refresh")
        if r:
            r.click(); page.wait_for_timeout(1500)
    return False
