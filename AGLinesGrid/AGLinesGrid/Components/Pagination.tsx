/**
 * @file Pagination.tsx
 * @description Component for rendering the pagination for the Lines.
 */

import * as React from "react";
import { Button, Text } from "@fluentui/react-components";
import {
    ArrowLeft16Regular,
    ArrowRight16Regular,
} from "@fluentui/react-icons";

interface CustomPaginationProps {
    currentPage: number;
    totalPages: number;
    onPageChange: (newPage: number) => void;
    getString: (s: string) => string;
}

export const CustomPagination: React.FC<CustomPaginationProps> = ({ currentPage, totalPages, onPageChange, getString }) => {
    const handlePrevClick = () => {
        if (currentPage > 1) {
            onPageChange(currentPage - 1);
        }
    };

    const handleNextClick = () => {
        if (currentPage < totalPages) {
            onPageChange(currentPage + 1);
        }
    };

    return (
        <div className="eg-pagination">
            <Button
                className="eg-pagination-button"
                disabled={currentPage === 1}
                onClick={handlePrevClick}
                icon={<ArrowLeft16Regular />}
                appearance="outline"
            ></Button>
            <Text>{getString("PaginationInfo").replace("{0}", currentPage.toString()).replace("{1}", totalPages.toString())}</Text>
            <Button
                className="eg-pagination-button"
                disabled={currentPage === totalPages}
                onClick={handleNextClick}
                icon={<ArrowRight16Regular />}
                appearance="outline"
            ></Button>
        </div>
    );
};
