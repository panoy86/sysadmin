from fastmcp import FastMCP #type: ignore

mcp = FastMCP()

@mcp.tool(name="greet", description="Say hello to someone")
def greet(name: str) -> str:
    return f"Hey {name}, how's it going?"

if __name__ == "__main__":
    try:
        mcp.run(transport="sse", host="127.0.0.1", port=8000)
    except Exception as e:
        import traceback
        print("An error occurred while running the MCP SSE server:")
        traceback.print_exc()