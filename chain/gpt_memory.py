from langchain_deepseek import ChatDeepSeek
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain
from langchain.callbacks.streaming_stdout import StreamingStdOutCallbackHandler
from langchain.memory import RedisChatMessageHistory, ConversationBufferMemory


class GptChain:
    template: str = """你是一位PPT生成专家，擅长根据用户的需求一步步生成合适的PPT内容
    
    {chat_history}
    Human: {human_input}
    Chatbot:"""
    openai_api_key: str = None
    session_id: str = None
    redis_url: str = None
    llm_chain: LLMChain = None
    message_history: RedisChatMessageHistory = None

    def __init__(self, openai_api_key, session_id, redis_url, openai_base_url):
        self.openai_api_key = openai_api_key
        self.session_id = session_id
        self.redis_url = redis_url
        self.openai_base_url = openai_base_url
        self.redis_llm_chain_factory()

    def redis_llm_chain_factory(self):
        """
        已经封装外部尽量不要调用此方法
        Returns:
        """
        message_history = RedisChatMessageHistory(
            url=self.redis_url, ttl=600, session_id=self.session_id
        )
        self.message_history = message_history
        memory = ConversationBufferMemory(
            memory_key="chat_history", chat_memory=message_history
        )
        prompt = PromptTemplate(
            input_variables=["chat_history", "human_input"], template=self.template)
        llm_chain = LLMChain(
            llm=ChatDeepSeek(model='deepseek-chat', api_key=self.openai_api_key, streaming=True,
                       callbacks=[StreamingStdOutCallbackHandler()]),
            prompt=prompt,
            verbose=True,
            memory=memory,
        )
        self.llm_chain = llm_chain

    def predict(self, question):
        return self.llm_chain.predict(human_input=question)

    def clear_redis(self):
        self.message_history.clear()


if __name__ == "__main__":
    chain = GptChain("you key", "1234", "you redis url")
    song = chain.predict(question="Write me a song about sparkling water.")
    # print(song)
