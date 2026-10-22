"""外部素材なしで作る、四枚の提案資料。座標の単位はポイント。"""

from slidekit import Deck, Math, Rect, Text


def build(context):
    deck = Deck(context)

    cover = deck.slide(
        id="cover",
        kind="cover",
        title="資料作成フローの提案",
        notes="資料の内容と表示を順に整える進め方を説明します。最後に、完成物を配布するまでの確認手順をまとめます。",
        seconds=20,
    )
    cover.text(
        id="author",
        text="作成者　slidekit example",
        area=Rect(96, 300, 768, 40),
        role="body",
        color="muted",
    )

    plan = deck.slide(
        id="plan",
        title="内容と表示の分離",
        notes="まず、左側の内容を決めます。目的と対象者を確認し、本文と図表を選び、根拠と説明をノートに残します。その後、右側の表示を整えます。最終的に配布する PPTX から作った画像で、文字や図形の状態を確認します。",
        seconds=50,
    )
    left, right = plan.content.columns(2, gap=48)
    for prefix, area, heading, labels in (
        ("content", left, "内容を決める", ("目的と対象者", "本文と図表", "根拠とノート")),
        ("appearance", right, "表示を整える", ("余白と文字", "図形と数式", "最終 PPTX の確認")),
    ):
        plan.text(id=f"{prefix}_heading", text=heading, area=Rect(area.x, 88, area.w, 40), role="heading")
        previous = None
        for key, y, label in zip(("first", "second", "third"), (156, 264, 372), labels):
            node = plan.box(
                id=f"{prefix}_{key}",
                area=Rect(area.x, y, area.w, 64),
                text=label,
                shape="rounded_rectangle",
                align="center",
                valign="middle",
            )
            if previous is not None:
                plan.arrow(
                    id=f"{prefix}_to_{key}",
                    start=previous.anchor("bottom"),
                    end=node.anchor("top"),
                    line_color="primary",
                    line_width_pt=1.5,
                )
            previous = node

    estimate = deck.slide(
        id="estimate",
        title="作業時間の見積もり",
        notes="説明のため、一枚当たりの作成時間と確認時間が一定と仮定します。作成時間を t_c、確認時間を t_r、枚数を n とします。合計時間 T は、一枚当たりの二つの時間の和に枚数を掛けた値です。数値の測定結果や、実際の作業時間を保証する式ではありません。図の二つの箱と矢印は一つのまとまりとして移動できます。数式を編集可能に保つため、PPTX 内では個別の図形として出力します。",
        seconds=65,
    )
    estimate.text(
        id="assumption",
        text="一枚当たりの作成時間と確認時間が一定と仮定",
        area=Rect(48, 90, 864, 40),
    )
    flow = estimate.group(id="work_cycle", area=Rect(96, 168, 768, 120))
    creation = flow.box(
        id="creation",
        text=[Text("作成\n"), Math(r"t_c")],
        area=Rect(0, 0, 300, 108),
        shape="rounded_rectangle",
        align="center",
        valign="middle",
    )
    review = flow.box(
        id="review",
        text=[Text("確認\n"), Math(r"t_r")],
        area=Rect(468, 0, 300, 108),
        shape="rounded_rectangle",
        align="center",
        valign="middle",
    )
    flow.arrow(
        id="to_review",
        start=creation.anchor("right"),
        end=review.anchor("left"),
        line_color="primary",
        line_width_pt=1.5,
    )
    estimate.equation(
        id="total_time",
        expression=r"T = n(t_c + t_r)",
        area=Rect(96, 332, 768, 72),
        align="center",
        valign="middle",
    )
    estimate.text(
        id="definitions",
        text=[Math("n"), Text("：枚数　"), Math("T"), Text("：合計時間")],
        area=Rect(96, 434, 768, 44),
        align="center",
    )

    summary = deck.slide(
        id="summary",
        title="導入の進め方",
        notes="まず少ない枚数で、依頼に必要な内容をそろえます。次に、生成時の診断を確認して修正し、一覧画像と各ページの画像で表示を確認します。数式の互換画像を付ける場合は、その画像を使う表示も確認します。本文、ノート、数式の意味がそろっていることを確認してから、完成物を決まった保存先へ反映します。",
        seconds=45,
    )
    summary.table(
        id="steps",
        headers=["段階", "確認すること"],
        rows=[
            ["内容をそろえる", "依頼の必須事項と説明の順序"],
            ["表示を整える", "文字、数式、図表の読みやすさ"],
            ["完成物を配布する", "本文、ノート、確認記録の一致"],
        ],
        row_keys=["prepare", "inspect", "deliver"],
        column_keys=["stage", "check"],
        area=Rect(48, 116, 864, 280),
        column_widths=[300, 564],
        row_heights=[64, 72, 72, 72],
        valign="middle",
        padding_pt=16,
    )
    summary.text(
        id="takeaway",
        text="最終 PPTX の表示と内容を確認してから配布",
        area=Rect(48, 438, 864, 44),
        role="heading",
    )
    return deck
