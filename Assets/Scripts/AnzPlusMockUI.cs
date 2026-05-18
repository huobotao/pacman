using UnityEngine;
using UnityEngine.UI;

public class AnzPlusMockUI : MonoBehaviour
{
    private readonly Color primaryBlue = new Color32(0x00, 0x7A, 0xD9, 0xFF);
    private readonly Color cardLight = new Color32(0xF3, 0xF8, 0xFD, 0xFF);

    private void Awake()
    {
        var canvas = BuildCanvas();
        BuildHeader(canvas.transform);
        BuildBalanceCard(canvas.transform);
        BuildQuickActions(canvas.transform);
        BuildTransactions(canvas.transform);
        BuildBottomNav(canvas.transform);
    }

    private Canvas BuildCanvas()
    {
        var canvasObject = new GameObject("ANZPlusCanvas", typeof(Canvas), typeof(CanvasScaler), typeof(GraphicRaycaster));
        var canvas = canvasObject.GetComponent<Canvas>();
        canvas.renderMode = RenderMode.ScreenSpaceOverlay;

        var scaler = canvasObject.GetComponent<CanvasScaler>();
        scaler.uiScaleMode = CanvasScaler.ScaleMode.ScaleWithScreenSize;
        scaler.referenceResolution = new Vector2(1170, 2532); // iPhone 13-like reference
        scaler.matchWidthOrHeight = 0.5f;

        var root = new GameObject("Background", typeof(RectTransform), typeof(Image));
        root.transform.SetParent(canvasObject.transform, false);
        var rootRect = root.GetComponent<RectTransform>();
        rootRect.anchorMin = Vector2.zero;
        rootRect.anchorMax = Vector2.one;
        rootRect.offsetMin = Vector2.zero;
        rootRect.offsetMax = Vector2.zero;
        root.GetComponent<Image>().color = Color.white;

        return canvas;
    }

    private void BuildHeader(Transform parent)
    {
        var header = CreatePanel("Header", parent, new Vector2(0.5f, 1f), new Vector2(1f, 1f), new Vector2(0, -160), new Vector2(-80, 220));
        header.color = primaryBlue;
        CreateText("Greeting", header.transform, "Hi, Alex", 64, FontStyle.Bold, TextAnchor.MiddleLeft, Color.white, new Vector2(70, -20), new Vector2(-120, 100));
        CreateText("Date", header.transform, "Monday 18 May", 32, FontStyle.Normal, TextAnchor.UpperLeft, new Color(1,1,1,0.9f), new Vector2(70, -95), new Vector2(-120, 60));
    }

    private void BuildBalanceCard(Transform parent)
    {
        var card = CreatePanel("BalanceCard", parent, new Vector2(0.5f, 1f), new Vector2(0.5f, 1f), new Vector2(0, -350), new Vector2(1000, 360));
        card.color = cardLight;

        CreateText("Label", card.transform, "Total balance", 34, FontStyle.Normal, TextAnchor.UpperLeft, Color.black, new Vector2(50, -50), new Vector2(900, 60));
        CreateText("Amount", card.transform, "$12,840.45", 72, FontStyle.Bold, TextAnchor.MiddleLeft, Color.black, new Vector2(50, -140), new Vector2(900, 120));
        CreateText("Subtitle", card.transform, "Spending + Savings", 30, FontStyle.Normal, TextAnchor.UpperLeft, new Color(0.35f,0.35f,0.35f), new Vector2(50, -230), new Vector2(900, 60));
    }

    private void BuildQuickActions(Transform parent)
    {
        var actions = CreatePanel("QuickActions", parent, new Vector2(0.5f, 1f), new Vector2(0.5f, 1f), new Vector2(0, -610), new Vector2(1000, 160));
        actions.color = Color.white;
        CreateRoundedButton("Pay", actions.transform, new Vector2(-320, 0));
        CreateRoundedButton("Transfer", actions.transform, new Vector2(0, 0));
        CreateRoundedButton("Bills", actions.transform, new Vector2(320, 0));
    }

    private void BuildTransactions(Transform parent)
    {
        var section = CreatePanel("Transactions", parent, new Vector2(0.5f, 1f), new Vector2(0.5f, 1f), new Vector2(0, -960), new Vector2(1000, 520));
        section.color = new Color(0.98f, 0.98f, 0.98f);
        CreateText("Title", section.transform, "Recent transactions", 36, FontStyle.Bold, TextAnchor.UpperLeft, Color.black, new Vector2(40, -40), new Vector2(920, 60));

        CreateTransactionRow(section.transform, "Woolworths", "Groceries", "-$84.20", -145);
        CreateTransactionRow(section.transform, "Uber", "Transport", "-$23.50", -255);
        CreateTransactionRow(section.transform, "Payroll", "Income", "+$2,800.00", -365, true);
    }

    private void BuildBottomNav(Transform parent)
    {
        var nav = CreatePanel("BottomNav", parent, new Vector2(0.5f, 0f), new Vector2(0.5f, 0f), new Vector2(0, 70), new Vector2(1000, 140));
        nav.color = Color.white;
        CreateText("HomeTab", nav.transform, "Home", 30, FontStyle.Bold, TextAnchor.MiddleCenter, primaryBlue, new Vector2(-300, 0), new Vector2(220, 100));
        CreateText("CardsTab", nav.transform, "Cards", 30, FontStyle.Normal, TextAnchor.MiddleCenter, Color.gray, new Vector2(0, 0), new Vector2(220, 100));
        CreateText("ProfileTab", nav.transform, "Profile", 30, FontStyle.Normal, TextAnchor.MiddleCenter, Color.gray, new Vector2(300, 0), new Vector2(220, 100));
    }

    private void CreateRoundedButton(string title, Transform parent, Vector2 position)
    {
        var button = CreatePanel(title + "Button", parent, new Vector2(0.5f, 0.5f), new Vector2(0.5f, 0.5f), position, new Vector2(280, 120));
        button.color = primaryBlue;
        CreateText(title + "Text", button.transform, title, 32, FontStyle.Bold, TextAnchor.MiddleCenter, Color.white, Vector2.zero, new Vector2(220, 80));
    }

    private void CreateTransactionRow(Transform parent, string merchant, string category, string amount, float top, bool positive = false)
    {
        CreateText(merchant + "Name", parent, merchant, 32, FontStyle.Bold, TextAnchor.UpperLeft, Color.black, new Vector2(40, top), new Vector2(500, 60));
        CreateText(merchant + "Category", parent, category, 26, FontStyle.Normal, TextAnchor.UpperLeft, Color.gray, new Vector2(40, top - 45), new Vector2(500, 50));
        CreateText(merchant + "Amount", parent, amount, 34, FontStyle.Bold, TextAnchor.MiddleRight, positive ? new Color(0.1f,0.6f,0.2f) : Color.black, new Vector2(540, top - 10), new Vector2(400, 60));
    }

    private Image CreatePanel(string name, Transform parent, Vector2 anchorMin, Vector2 anchorMax, Vector2 anchoredPos, Vector2 size)
    {
        var panel = new GameObject(name, typeof(RectTransform), typeof(Image));
        panel.transform.SetParent(parent, false);
        var rect = panel.GetComponent<RectTransform>();
        rect.anchorMin = anchorMin;
        rect.anchorMax = anchorMax;
        rect.anchoredPosition = anchoredPos;
        rect.sizeDelta = size;
        return panel.GetComponent<Image>();
    }

    private Text CreateText(string name, Transform parent, string value, int fontSize, FontStyle style, TextAnchor align, Color color, Vector2 anchoredPos, Vector2 size)
    {
        var textObject = new GameObject(name, typeof(RectTransform), typeof(Text));
        textObject.transform.SetParent(parent, false);
        var rect = textObject.GetComponent<RectTransform>();
        rect.anchorMin = new Vector2(0, 1);
        rect.anchorMax = new Vector2(0, 1);
        rect.pivot = new Vector2(0, 1);
        rect.anchoredPosition = anchoredPos;
        rect.sizeDelta = size;

        var text = textObject.GetComponent<Text>();
        text.font = Resources.GetBuiltinResource<Font>("Arial.ttf");
        text.text = value;
        text.fontSize = fontSize;
        text.fontStyle = style;
        text.alignment = align;
        text.color = color;
        return text;
    }
}
