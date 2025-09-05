# chatbot.py
from model import ai

def chatbot(summary: str, question: str) -> str:
    """
    Chatbot function to answer a single question based on a CSV summary.

    Args:
        summary (str): Summary/context of the CSV file.
        question (str): User's question about the data.

    Returns:
        str: AI-generated answer or an error message.
    """
    prompt = f"""
You are a helpful, factual data analyst assistant. Use only the information contained in the dataset summary below to answer the user's question.
If the question cannot be answered from the summary, say you don't have enough information and suggest what data would be helpful.

Dataset summary:
{summary}

Question:
{question}

Answer concisely and clearly, and if helpful include a short next-step suggestion (one line).
"""
    try:
        response = ai(prompt)
        return response.strip()
    except Exception as e:
        return f"Error generating answer: {e}"

def chatbot_loop(summary: str) -> None:
    """
    Start an interactive chatbot loop that uses the dataset summary as context.
    The loop continues until the user types /exit, exit, quit, or presses Ctrl+C.

    Args:
        summary (str): Summary/context of the CSV file.

    Returns:
        None
    """
    print("\n🤖 Chatbot mode — interactive Q&A based on the dataset summary.")
    print("Type /exit to quit, or press Ctrl+C.")
    print("=" * 60)
    if summary:
        print("Dataset summary (used as context):")
        print(summary)
        print("-" * 60)

    try:
        while True:
            try:
                user_q = input("\nYou: ").strip()
            except KeyboardInterrupt:
                print("\n👋 Exiting chatbot.")
                break

            if not user_q:
                continue
            if user_q.lower() in ("/exit", "exit", "quit"):
                print("👋 Exiting chatbot.")
                break

            answer = chatbot(summary, user_q)
            print("\n🤖", answer)

    except Exception as e:
        # Catch unexpected exceptions to avoid breaking host programs
        print(f"Chatbot encountered an error and will exit: {e}")
