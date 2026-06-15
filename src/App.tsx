import {
  ArrowUpRight,
  Boxes,
  CheckCircle2,
  Database,
  FolderGit2,
  Image,
  MessageSquareText,
  Network,
  ShieldCheck
} from "lucide-react";

const localCrmUrl = "http://127.0.0.1:8770/?view=crm&standalone=1";

const modules = [
  { label: "업무관리 CRM", detail: "대시보드, 내 업무, 업무보드, 거래처, 메신저 연동", icon: Boxes },
  { label: "직원 대시보드", detail: "조직도형 직원 현황과 역할 중심 보기", icon: Network },
  { label: "사내 메신저", detail: "직원 간 메시지 저장과 조회 API 포함", icon: MessageSquareText },
  { label: "운영 데이터", detail: "DB, 토큰, 백업 파일은 GitHub 업로드 제외", icon: ShieldCheck }
];

const projectFiles = [
  "package.json",
  "src/",
  "vite.config.ts",
  "tsconfig.json",
  "_workhub_zip_inspect/scripts/workhub_delivery_app.py"
];

const mockups = [
  {
    title: "통합 대시보드",
    src: "/mockups/workhub-crm-integrated-dashboard.png"
  },
  {
    title: "고객 상세",
    src: "/mockups/workhub-crm-customer-detail.png"
  },
  {
    title: "업무보드",
    src: "/mockups/crm_board_collapsible_final_104900.png"
  }
];

export function App() {
  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div className="brand-block">
          <div className="brand-mark">W</div>
          <div>
            <p className="brand-name">Workhub CRM</p>
            <p className="brand-subtitle">Startup Operations Suite</p>
          </div>
        </div>

        <nav className="nav-list" aria-label="프로젝트 메뉴">
          <a href="#overview">개요</a>
          <a href="#modules">모듈</a>
          <a href="#mockups">목업</a>
          <a href="#repository">저장소 구성</a>
        </nav>

        <a className="local-link" href={localCrmUrl} target="_blank" rel="noreferrer">
          로컬 CRM 열기
          <ArrowUpRight size={16} />
        </a>
      </aside>

      <section className="content">
        <header id="overview" className="hero">
          <div>
            <h1>스타트업 업무 관리 CRM 프로젝트</h1>
            <p>
              발주, 수출입, 거래처, 직원 조직도, 사내 메신저를 하나의 운영 화면으로 묶은
              Workhub 프로젝트야. 이 프론트엔드 셸은 GitHub와 배포 도구가 프로젝트 구조를
              바로 읽을 수 있게 만든 Vite 기반 진입점이야.
            </p>
          </div>
          <div className="hero-card" aria-label="프로젝트 상태">
            <CheckCircle2 size={22} />
            <strong>GitHub-ready</strong>
            <span>Vite, React, TypeScript 구조 추가 완료</span>
          </div>
        </header>

        <section id="modules" className="section">
          <div className="section-heading">
            <h2>주요 모듈</h2>
            <p>기존 파이썬 CRM 앱과 새 프론트엔드 구조가 함께 관리돼.</p>
          </div>
          <div className="module-grid">
            {modules.map((item) => {
              const Icon = item.icon;
              return (
                <article className="module-card" key={item.label}>
                  <Icon size={22} />
                  <h3>{item.label}</h3>
                  <p>{item.detail}</p>
                </article>
              );
            })}
          </div>
        </section>

        <section id="mockups" className="section">
          <div className="section-heading">
            <h2>목업 미리보기</h2>
            <p>루트 저장소에 포함된 디자인 확인용 이미지들이야.</p>
          </div>
          <div className="mockup-grid">
            {mockups.map((mockup) => (
              <figure className="mockup-card" key={mockup.src}>
                <img src={mockup.src} alt={`${mockup.title} 목업`} />
                <figcaption>
                  <Image size={16} />
                  {mockup.title}
                </figcaption>
              </figure>
            ))}
          </div>
        </section>

        <section id="repository" className="section repository-panel">
          <div>
            <div className="section-heading compact">
              <h2>저장소 구성</h2>
              <p>GitHub에는 프론트엔드 진입점과 실제 CRM 앱 파일이 같이 올라가.</p>
            </div>
            <ul className="file-list">
              {projectFiles.map((file) => (
                <li key={file}>
                  <FolderGit2 size={16} />
                  <code>{file}</code>
                </li>
              ))}
            </ul>
          </div>
          <div className="data-note">
            <Database size={22} />
            <strong>민감 파일 제외</strong>
            <p>운영 DB, 웹훅 토큰, 백업 zip, 캐시는 `.gitignore`로 제외했어.</p>
          </div>
        </section>
      </section>
    </main>
  );
}
