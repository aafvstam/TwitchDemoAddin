import * as React from "react";
import { NavLink } from "react-router-dom";

export interface HeaderProps {
  title: string;
  logo?: string;
  message?: string;
}

function Header(props: HeaderProps) {
  const title = props.title;
  const activeStyle = { color: "orange" };

  return (
    <section className="ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500">
      <h1 className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">{title}</h1>
      <nav>
        <NavLink activeStyle={activeStyle} exact to="/">
          Home
        </NavLink>
        {" | "}
        <NavLink activeStyle={activeStyle} to="/watermark">
          Watermark
        </NavLink>
        {" | "}
        <NavLink activeStyle={activeStyle} to="/about">
          About
        </NavLink>
      </nav>
    </section>
  );
}

export default Header;
