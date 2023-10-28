import * as React from "react";
import styles from "./CustomControls.module.scss";
import type { ICustomControlsProps } from "./ICustomControlsProps";
import { WebPartTitle } from "../../../common/components/webPartTitle";

const CustomControls: React.FC<ICustomControlsProps> = ({
  headerProps,
  hasTeamsContext,
  listId,
}) => {
  return (
    <section
      className={`${styles.webPartRoot} ${hasTeamsContext ? styles.teams : ""}`}
    >
      <div className={styles.container}>
        <WebPartTitle
          title={headerProps.title}
          description={headerProps.description}
        />

        <h5>Selected list: {listId}</h5>
      </div>
    </section>
  );
};

export default CustomControls;
