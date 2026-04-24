use pbo_server::{router, AppState};

#[tokio::main]
async fn main() {
    let state = AppState::from_workspace_root();
    let app = router(state);

    let listener = tokio::net::TcpListener::bind("0.0.0.0:3001").await.unwrap();
    println!("PolarBearOffice server listening on http://localhost:3001");
    axum::serve(listener, app).await.unwrap();
}
