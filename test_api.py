import requests

# ضع مفتاحك الجديد هنا مباشرة (للتجربة فقط)
api_key = "sk-nry-rLLo3nMR2k9OPje9wymif0XLj5KeKkXbNn9aYFGV6aY" 
api_url = "https://router.bynara.id/v1/chat/completions"

print(f"--- بدء الاختبار المباشر ---")

headers = {
    "Authorization": f"Bearer {api_key}",
    "Content-Type": "application/json"
}

data = {
    "model": "claude-sonnet-4-5",
    "messages": [{"role": "user", "content": "Hello"}]
}

try:
    response = requests.post(api_url, headers=headers, json=data)
    print(f"النتيجة: {response.status_code}")
    if response.status_code == 200:
        print("✅ نجاح! الاتصال يعمل والمفتاح صحيح.")
    else:
        print(f"❌ فشل الاتصال: {response.text}")
except Exception as e:
    print(f"حدث خطأ: {e}")