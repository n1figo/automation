# Welcome to your Lovable project

pyinstaller --onefile --windowed pdf_analyzer.spec

pdf_analyzer/
├── main.py           # Tkinter 애플리케이션 코드
├── analyzer/         # 분석 로직 분리
│   ├── __init__.py
│   ├── pdf_parser.py # PDF 파싱 함수
│   ├── type_finder.py # 종별 찾기 함수
│   └── highlight_analyzer.py # 강조색 분석 클래스
├── utils/            # 유틸리티 함수
│   ├── __init__.py
│   └── excel_writer.py # Excel 생성 함수
└── resources/        # 필요한 리소스
    └── icon.ico      # 애플리케이션 아이콘

## Project info

**URL**: https://lovable.dev/projects/484f858e-a17d-42a5-97d3-5164c4a5af47

## How can I edit this code?

There are several ways of editing your application.

**Use Lovable**

Simply visit the [Lovable Project](https://lovable.dev/projects/484f858e-a17d-42a5-97d3-5164c4a5af47) and start prompting.

Changes made via Lovable will be committed automatically to this repo.

**Use your preferred IDE**

If you want to work locally using your own IDE, you can clone this repo and push changes. Pushed changes will also be reflected in Lovable.

The only requirement is having Node.js & npm installed - [install with nvm](https://github.com/nvm-sh/nvm#installing-and-updating)

Follow these steps:

```sh
# Step 1: Clone the repository using the project's Git URL.
git clone <YOUR_GIT_URL>

# Step 2: Navigate to the project directory.
cd <YOUR_PROJECT_NAME>

# Step 3: Install the necessary dependencies.
npm i

# Step 4: Start the development server with auto-reloading and an instant preview.
npm run dev
```

**Edit a file directly in GitHub**

- Navigate to the desired file(s).
- Click the "Edit" button (pencil icon) at the top right of the file view.
- Make your changes and commit the changes.

**Use GitHub Codespaces**

- Navigate to the main page of your repository.
- Click on the "Code" button (green button) near the top right.
- Select the "Codespaces" tab.
- Click on "New codespace" to launch a new Codespace environment.
- Edit files directly within the Codespace and commit and push your changes once you're done.

## What technologies are used for this project?

This project is built with .

- Vite
- TypeScript
- React
- shadcn-ui
- Tailwind CSS

## How can I deploy this project?

Simply open [Lovable](https://lovable.dev/projects/484f858e-a17d-42a5-97d3-5164c4a5af47) and click on Share -> Publish.

## I want to use a custom domain - is that possible?

We don't support custom domains (yet). If you want to deploy your project under your own domain then we recommend using Netlify. Visit our docs for more details: [Custom domains](https://docs.lovable.dev/tips-tricks/custom-domain/)
