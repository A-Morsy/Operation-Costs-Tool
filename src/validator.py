from langchain.chains import create_stuff_documents_chain, create_retrieval_chain
from langchain_core.prompts import ChatPromptTemplate
from langchain_ollama import ChatOllama
from langchain_community.embeddings import OllamaEmbeddings
from langchain_community.vectorstores import FAISS

# 1. Prepare retriever (as before)
retriever = vectorstore.as_retriever()

# 2. Craft prompts
system_prompt = (
    "Use the following context to answer the question. "
    "If you're unsure, respond with 'I don't know'. "
    "Context: {context}"
)
prompt_template = ChatPromptTemplate.from_messages([
    ("system", system_prompt),
    ("human", "{input}")
])

# 3. Create a chain to combine documents
stuff_chain = create_stuff_documents_chain(llm=ChatOllama(model="mistral:latest", temperature=0.3),
                                           prompt=prompt_template)

# 4. Build the retrieval chain
chain = create_retrieval_chain(retriever, stuff_chain)

# 5. Run your query
result = chain.invoke({"input": "Which employee has the lowest costs for each operation?"})
print(result["output"])
