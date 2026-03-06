import asyncio
import json
import sys, os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from aiocua import CuaOperator


async def main():
    operator = CuaOperator()

    # Get the full desktop accessibility tree
    tree_json = await operator.axtree()
    tree = json.loads(tree_json)

    await asyncio.sleep(1)

    # Find a Chrome window node
    chrome_node_id = None

    def find_chrome(node):
        nonlocal chrome_node_id
        if chrome_node_id:
            return
        name = node.get("name", "")
        role = node.get("role", "")
        if role == "window" and "chrome" in name.lower():
            chrome_node_id = node["id"]
            return
        for child in node.get("children", []):
            find_chrome(child)

    find_chrome(tree)

    if not chrome_node_id:
        print("Error: Chrome window not found.")
        return

    # Focus Chrome
    await operator.ax_focus(chrome_node_id)
    await asyncio.sleep(1)

    # Focus address bar and navigate
    await operator.ax_keypress(["CONTROL", "l"])
    await asyncio.sleep(0.5)
    await operator.ax_type("http://www.google.com")
    await asyncio.sleep(0.5)
    await operator.ax_keypress(["RETURN"])
    await asyncio.sleep(2)

    # Print Chrome's subtree
    subtree = await operator.axtree(root_node_id=chrome_node_id)
    print(subtree)


if __name__ == "__main__":
    asyncio.run(main())
