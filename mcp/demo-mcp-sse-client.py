
import asyncio
from fastmcp import Client #type: ignore

async def main():
    async with Client("http://localhost:8000/sse") as client:
        print(await client.list_tools())
        result = await client.call_tool("greet", {"name": "Alice"})
        print("Result from greet tool:")
        #print(result)
        #print(result.content)
        print(result.content[0].text)

asyncio.run(main())
