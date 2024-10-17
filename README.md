## Upute za generiranje Excel tablice postova

#### 1. Instaliraj python (verzija u projektu 3.9.13)
#### 2. Kreiraj virtualno okruženje
```sh
python -m venv myVenv
```
#### 3. Aktiviraj virtualno okruženje

```sh
source myVenv/Scripts/activate
```
#### 4. Instaliraj potrebne pakete (iz requirements.txt)

```sh
pip install -r requirements.txt
```

#### 5. Preuzmni xml file kolkcije postova iz wordpress cms-a
* [link na cms](https://admin.uplift.hr/wp-admin/)
* Tools → Posts → Download Export File button

#### 6. Preimenuj ga u "posts.xml" i stavi u isti folder gdje je i skrpta "extractor.py"

#### 7. U terminal upiši
```sh
python extractor.py
```

#### 8. File kreiran DONE!