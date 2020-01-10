import "bootstrap/dist/css/bootstrap.css";
import * as React from "react";
import { HashRouter as Router, NavLink } from "react-router-dom";

export interface HeaderProps {
  title: string;
  logo?: string;
  message?: string;
}

function Header(props: HeaderProps) {
  const title = props.title;
  const logo = props.logo;
  const activeStyle = { color: "orange" };

  return (
    <section className="ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500">
      <img width="90" height="90" src={logo} alt={title} title={title} />
      <h1 className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">{title}</h1>
      <Router>
        <nav className="navbar navbar-expand-md navbar-dark bg-transparent">
          <ul className="navbar-nav mr-auto">
            <li className="nav-item">
              <NavLink activeStyle={activeStyle} exact to="/">
                Home
              </NavLink>
            </li>
            <li className="nav-item">
              <NavLink activeStyle={activeStyle} to="/Watermark">
                Watermark
              </NavLink>
            </li>
            <li className="nav-item">
              <NavLink activeStyle={activeStyle} to="/About">
                About
              </NavLink>
            </li>
          </ul>
        </nav>
      </Router>
    </section>
  );
}

export default Header;
