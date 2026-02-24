"""Chart Extension objects and related items."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx.shared import ElementProxy
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.chart.chartex import CT_ChartSpace, CT_Series, CT_Axis
    from pptx.parts.chartex import ChartExPart


class ChartEx(ElementProxy):
    """Chart extension object.
    
    Corresponds to the ``<cx:chartSpace>`` element that is the root of a chart extension part.
    """
    
    _chartspace: CT_ChartSpace
    
    def __init__(self, chartSpace: CT_ChartSpace, chart_part: ChartExPart):
        super(ChartEx, self).__init__(chartSpace, chart_part)
        self._chartspace = chartSpace
    
    @property
    def chart_title(self) -> str | None:
        """The title of this chart as a string, or None if there is no title.
        
        Assigning a string value sets the title to that value. Assigning None causes
        any title to be deleted.
        """
        title = self._chart.title
        if title is None:
            return None
        
        title_text = title.find(".//cx:txData/cx:v", namespaces={"cx": "http://schemas.microsoft.com/office/drawing/2014/chartex"})
        if title_text is None:
            return None
        
        return title_text.text
    
    @property
    def chart_type(self) -> str | None:
        """The chart type of this chart extension as a string.
        
        Possible values include:
        * waterfall
        * sunburst
        * treemap
        * funnel
        * boxWhisker
        * clusteredColumn
        * paretoLine
        * regionMap
        
        Returns None if the chart type cannot be determined.
        """
        plotAreaRegion = self._chart.plotArea.plotAreaRegion
        series = plotAreaRegion.find(".//cx:series", namespaces={"cx": "http://schemas.microsoft.com/office/drawing/2014/chartex"})
        if series is None:
            return None
        
        return series.get("layoutId")
        
    @property
    def has_legend(self) -> bool:
        """True if this chart has a legend, False otherwise."""
        return self._chart.legend is not None
    
    @lazyproperty
    def legend(self) -> Legend:
        """
        The |Legend| instance for this chart, providing properties and methods on the
        chart legend.
        """
        legend = self._chart.get_or_add_legend()
        return Legend(legend, self)
    
    @property
    def series(self) -> list[Series]:
        """A sequence of |Series| objects representing the series in this chart."""
        series_elements = self._chart.plotArea.plotAreaRegion.findall(".//cx:series", namespaces={"cx": "http://schemas.microsoft.com/office/drawing/2014/chartex"})
        return [Series(series, self) for series in series_elements]
    
    @property
    def axes(self) -> list[Axis]:
        """A sequence of |Axis| objects representing the axes in this chart."""
        axis_elements = self._chart.plotArea.findall(".//cx:axis", namespaces={"cx": "http://schemas.microsoft.com/office/drawing/2014/chartex"})
        return [Axis(axis, self) for axis in axis_elements]
    
    @property
    def _chart(self):
        """The ``<cx:chart>`` element in this chart."""
        return self._chartspace.chart


class Legend(ElementProxy):
    """Chart legend object.
    
    Corresponds to the ``<cx:legend>`` element in chartex.
    """
    
    @property
    def position(self) -> str | None:
        """
        Return the position of the legend as a string, or None if the position
        is not specified. Valid values are 'l', 't', 'r', 'b' indicating left,
        top, right, or bottom.
        """
        pos = self._element.get("pos")
        if pos is None:
            return None
        return pos
    
    @position.setter
    def position(self, value):
        """
        Set the position of the legend to one of 'l', 't', 'r', or 'b'.
        """
        valid_positions = {'l', 't', 'r', 'b'}
        if value not in valid_positions:
            raise ValueError(f"position must be one of {', '.join(valid_positions)}")
        self._element.set("pos", value)
    
    @property
    def include_in_layout(self) -> bool:
        """
        Return True if the legend's position is affected by the chart layout,
        False otherwise.
        """
        overlay = self._element.get("overlay")
        if overlay is None:
            return True
        return overlay == "0"
    
    @include_in_layout.setter
    def include_in_layout(self, value):
        """
        Set whether the legend's position is affected by the chart layout.
        """
        self._element.set("overlay", "0" if value else "1")


class Series(ElementProxy):
    """Chart series object.
    
    Corresponds to the ``<cx:series>`` element in chartex.
    """
    
    _series: CT_Series
    
    def __init__(self, series: CT_Series, parent: ChartEx):
        super(Series, self).__init__(series, parent)
        self._series = series
    
    @property
    def name(self) -> str | None:
        """The name of this series, or None if it has no name."""
        tx = self._series.tx
        if tx is None:
            return None
        
        tx_text = tx.find(".//cx:v", namespaces={"cx": "http://schemas.microsoft.com/office/drawing/2014/chartex"})
        if tx_text is None:
            return None
            
        return tx_text.text
    
    @property
    def is_visible(self) -> bool:
        """True if this series is visible, False otherwise."""
        hidden = self._series.get("hidden")
        if hidden is None:
            return True
        return hidden == "0"
    
    @is_visible.setter
    def is_visible(self, value):
        """Set whether this series is visible."""
        self._series.set("hidden", "0" if value else "1")
    
    @property
    def values(self) -> list[float | None]:
        """The data values for this series."""
        # This is a simplified implementation that doesn't handle real data retrieval
        # In a real implementation, you would need to retrieve data from the Excel workbook
        # based on the dataId and other information in the series
        return []


class Axis(ElementProxy):
    """Chart axis object.
    
    Corresponds to the ``<cx:axis>`` element in chartex.
    """
    
    _axis: CT_Axis
    
    def __init__(self, axis: CT_Axis, parent: ChartEx):
        super(Axis, self).__init__(axis, parent)
        self._axis = axis
    
    @property
    def id(self) -> int:
        """The id of this axis."""
        return int(self._axis.get("id"))
    
    @property
    def is_visible(self) -> bool:
        """True if this axis is visible, False otherwise."""
        hidden = self._axis.get("hidden")
        if hidden is None:
            return True
        return hidden == "0"
    
    @is_visible.setter
    def is_visible(self, value):
        """Set whether this axis is visible."""
        self._axis.set("hidden", "0" if value else "1")
    
    @property
    def has_major_gridlines(self) -> bool:
        """True if this axis has major gridlines, False otherwise."""
        return self._axis.majorGridlines is not None
    
    @property
    def has_minor_gridlines(self) -> bool:
        """True if this axis has minor gridlines, False otherwise."""
        return self._axis.minorGridlines is not None
    
    @property
    def title(self) -> str | None:
        """The title of this axis, or None if it has no title."""
        title = self._axis.title
        if title is None:
            return None
        
        title_text = title.find(".//cx:v", namespaces={"cx": "http://schemas.microsoft.com/office/drawing/2014/chartex"})
        if title_text is None:
            return None
            
        return title_text.text