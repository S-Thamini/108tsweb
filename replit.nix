{ pkgs }: {
  deps = [
    pkgs.python311Full
    pkgs.python311Packages.pip
    pkgs.python311Packages.pandas
    pkgs.python311Packages.fastapi
    pkgs.python311Packages.uvicorn
    pkgs.python311Packages.openpyxl
    pkgs.python311Packages.requests
    pkgs.python311Packages.python-multipart
  ];
}
