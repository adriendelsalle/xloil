#ifndef NOMINMAX
#define NOMINMAX
#endif

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#include <windows.h>

// Add-in designer C:\\Program Files\\Common Files\\Designer\\MSADDNDR.OLB
#import "libid:AC0714F2-3D04-11D1-AE7D-00A0C90F26F4"

class D : public AddInDesignerObjects::_IDTExtensibility2 {};

int main() {
  AddInDesignerObjects::IDTExtensibility2 a;
  return 0;
}