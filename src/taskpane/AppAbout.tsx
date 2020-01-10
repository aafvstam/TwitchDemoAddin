import * as React from "react";
import Header from "./components/common/Header";
import AboutPage from "./pages/AboutPage";

export interface AppProps {
  title: string;
}

export interface AppState {
}

export default class App extends React.Component<AppProps, AppState> {
  render() {
    const { title } = this.props;

    return (
      <div className="ms-welcome">
        <Header logo="assets/profile300x300.png" title={title} message="About" />
        <AboutPage/>
      </div>
    );
  }
}
