# GitHub 連携手順

このリポジトリを [sogakouki1022-sogasoga](https://github.com/sogakouki1022-sogasoga) の GitHub にプッシュする手順です。

## 1. GitHub で新しいリポジトリを作成

1. https://github.com/new を開く
2. **Repository name** を入力（例: `mypj`）
3. **Public** または **Private** を選択
4. **「Add a README file」はチェックしない**（既にローカルにコミット済みのため）
5. **Create repository** をクリック

## 2. リモートを追加してプッシュ（PowerShell）

リポジトリ名を `mypj` にした場合の例です。

```powershell
cd "c:\repo\mypj"

# リモート追加（リポジトリ名は実際に作成した名前に合わせる）
git remote add origin https://github.com/sogakouki1022-sogasoga/mypj.git

# 初回プッシュ
git push -u origin main
```

リポジトリ名を別の名前にした場合は、URL の `mypj` の部分をその名前に変更してください。

## 3. 認証について

- **HTTPS** の場合: プッシュ時に GitHub のユーザー名とパスワード（または Personal Access Token）の入力が求められます
- **SSH** を使う場合:
  ```powershell
  git remote add origin git@github.com:sogakouki1022-sogasoga/mypj.git
  ```

## 4. 今後の運用

- 変更をコミット: `git add -A` → `git commit -m "メッセージ"`
- プッシュ: `git push`
