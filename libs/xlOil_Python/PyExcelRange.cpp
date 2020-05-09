#include "InjectedModule.h"
#include "PyHelpers.h"
#include "BasicTypes.h"
#include <xlOil/ExcelRange.h>

using std::shared_ptr;
using std::vector;
namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    namespace
    {
      // Works like the Range.Range function in VBA which is 1-based and
      // includes the right hand end-point
      inline auto subRange(const Range& r,
        int fromR, int fromC,
        int* toR = 0, int* toC = 0, 
        size_t* nRows = 0, size_t* nCols = 0)
      {
        if (!(toR || nRows))
          XLO_THROW("Must specify end row or number of rows");
        if (!(toC || nCols))
          XLO_THROW("Must specify end column or number of columns");
        --fromR; --fromC; // Correct for 1-based indexing
        return r.range(fromR, fromC,
          toR ? *toR : fromR + *nRows,
          toC ? *toC : fromC + *nCols);
      }

      // Works like the Range.Cell function in VBA which is 1-based
      inline auto rangeCell(const Range& r, int row, int col)
      {
        return r.cell(row - 1, col - 1);
      }
      auto convertExcelObj(ExcelObj&& val)
      {
        return PySteal<>(PyFromExcel<PyFromAny<>>()(val));
      }
      auto rangeGetValue(const Range& r)
      {
        return convertExcelObj(r.value());
      }
      void rangeSetValue(Range& r, const py::object& pyval)
      {
        const auto val = FromPyObj()(pyval.ptr());
        // Release gil when setting values in as this may trigger calcs 
        // which call back into other python functions.
        py::gil_scoped_release lose_gil;
        r = val;
      };

      void rangeClear(Range& r)
      {
        py::gil_scoped_release lose_gil;
        r.clear();
      }

      py::object getItem(const Range& range, pybind11::tuple loc)
      {
        if (loc.size() != 2)
          XLO_THROW("Expecting tuple of size 2");
        auto r = loc[0];
        auto c = loc[1];
        size_t fromRow, fromCol, toRow, toCol, step = 1;
        if (r.is_none())
        {
          fromRow = 0;
          toRow = range.nRows();
        }
        else if (PySlice_Check(r.ptr()))
        {
          size_t sliceLength;
          r.cast<py::slice>().compute(range.nRows(), &fromRow, &toRow, &step, &sliceLength);
        }
        else
        {
          fromRow = r.cast<size_t>();
          toRow = fromRow + 1;
        }

        if (r.is_none())
        {
          fromCol = 0;
          toCol = range.nRows();
        }
        else if (PySlice_Check(c.ptr()))
        {
          size_t sliceLength;
          c.cast<py::slice>().compute(range.nCols(), &fromCol, &toCol, &step, &sliceLength);
        }
        else
        {
          fromCol = c.cast<size_t>();
          // Check for single element access
          if (fromRow == toRow + 1)
            return convertExcelObj(range.value((int)fromRow, (int)fromCol));
          toCol = fromCol + 1;
        }

        if (step != 1)
          XLO_THROW("Slices step size must be 1");
        
        return py::cast(range.range(fromRow, fromCol, toRow, toCol));
      }

      static int theBinder = addBinder([](pybind11::module& mod)
      {
        // Bind Range class from xloil::ExcelRange
        auto rType = py::class_<Range>(mod, "Range")
          .def(py::init([](const wchar_t* x) { return newRange(x); }),
            py::arg("address"))
          .def("range", subRange,
            py::arg("from_row"), 
            py::arg("from_col"),
            py::arg("to_row") = nullptr,
            py::arg("to_col") = nullptr,
            py::arg("num_rows") = -1, 
            py::arg("num_cols") = -1)
          .def("cells", rangeCell,
            py::arg("row"), 
            py::arg("col"))
          .def_property("value",
            rangeGetValue,
            rangeSetValue,
            py::return_value_policy::automatic)
          .def("set", rangeSetValue)
          .def("clear", rangeClear)
          .def("address", [](const Range& r, bool local) { return r.address(local); },
            py::arg("local") = false)
          .def_property_readonly("nrows", [](const Range& r) { return r.nRows(); })
          .def_property_readonly("ncols", [](const Range& r) { return r.nCols(); })
          .def_property_readonly("shape", 
            [](const Range& r)
            {
              return std::make_pair(r.nRows(), r.nCols());
            });

      }, 99);
    }
  }
}