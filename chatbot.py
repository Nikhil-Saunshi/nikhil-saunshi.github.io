import openai
import time

openai.api_key = "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" # replace with your API key

def generate_response(prompt):
    response = openai.Completion.create(
        engine="davinci",
        prompt=prompt,
        temperature=0.7,
        max_tokens=10,
        top_p=1,
        n=1,
        stop="\n",
        presence_penalty=0,
        frequency_penalty=0
    )
    message = response.choices[0].text.strip()
    return message

while True:
    user_input = input("You: ")
    if user_input.lower() == "bye":
        print("Lamebot: Goodbye!")
        break
    prompt = f"User: {user_input}\nLamebot:"
    response = generate_response(prompt)
    # Needs data cleaning here
    print("Lamebot:", response)
