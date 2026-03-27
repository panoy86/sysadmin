# Demo using STDIO

from fastmcp import FastMCP #type: ignore
mcp = FastMCP("SimpleServer")

@mcp.tool()
def add(a: int, b: int) -> int:
    """Add two numbers"""
    return a + b

if __name__ == "__main__":
    mcp.run()