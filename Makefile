icon_file=document_distributor.ico
entry_point=document_distributor/_gui.py
app_name=ddist

.PHONY: build-linux
build-linux:
	pyinstaller --icon $(icon_file) -n $(app_name) $(entry_point)

.PHONY: build-win
build-win:
	pyinstaller --icon $(icon_file) -n $(app_name) -w $(entry_point)
