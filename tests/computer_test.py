from aiocua import CuaOperator
from aiocua.contracts.computer import MonitorMetadata
import pytest


@pytest.mark.asyncio
async def test_dimensions_screenshot():
    operator = CuaOperator()
    dimensions = await operator.dimensions()
    assert isinstance(dimensions, tuple)
    assert len(dimensions) == 2
    assert isinstance(dimensions[0], int)
    assert isinstance(dimensions[1], int)
    assert dimensions[0] > 0
    assert dimensions[1] > 0

    screenshot = await operator.screenshot()
    assert isinstance(screenshot, str)

    import base64
    from io import BytesIO
    from PIL import Image

    img = Image.open(BytesIO(base64.b64decode(screenshot)))
    assert img.size == dimensions


@pytest.mark.asyncio
async def test_monitors():
    operator = CuaOperator()
    result = await operator.monitors()
    assert isinstance(result, list)
    assert len(result) > 0

    for m in result:
        assert isinstance(m, MonitorMetadata)
        assert isinstance(m.id, int)
        assert isinstance(m.width, int)
        assert isinstance(m.height, int)
        assert m.width > 0
        assert m.height > 0
        assert isinstance(m.scale, float)
        assert m.scale >= 1.0
        assert isinstance(m.primary, bool)

    primaries = [m for m in result if m.primary]
    assert len(primaries) == 1

    dimensions = await operator.dimensions()
    primary = primaries[0]
    assert (primary.width, primary.height) == dimensions
