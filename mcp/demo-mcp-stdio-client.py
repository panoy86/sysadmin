# Demo using STDIO

import asyncio
from fastmcp import Client #type: ignore
async def main():
    async with Client("demo-mcp-stdio-server.py") as client:
        tools = await client.list_tools()
        print("Available tools:", tools)
        result = await client.call_tool("add", {"a": 5, "b": 7})
        print("Result:", result.content[0].text)
asyncio.run(main())