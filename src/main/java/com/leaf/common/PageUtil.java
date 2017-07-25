package com.leaf.common;


import java.io.Serializable;
import java.util.List;

/**
 * 作者 叶舟
 * 时间 2017年5月12日下午5:01:38
 */
public class PageUtil<T> implements Serializable{

	private static final long serialVersionUID = 3509748940303200410L;

	/**

	 * 对象信息

	 */
	private List<T> pageResults;

	/**

	 * 总条数

	 */
	private Integer totalCount;

	/**

	 * 当前页

	 */
	private Integer pageNo;

	/**

	 * 每页显示的记录数

	 */
	private Integer pageSize;

	/**

	 * 总页码

	 */
	private Integer totalPage;


	/**

	 * 开始行数

	 */
	private Integer startRow;
	
	public PageUtil(){
		this.setPageSize(10);
	}
	
	public PageUtil(int pageSize, int tolalCount) {
		this.setPageSize(pageSize);
		this.setTotalCount(tolalCount);
	}
	
	public PageUtil(int pageSize, int tolalCount,List<T> list) {
		this.setPageSize(pageSize);
		this.setTotalCount(tolalCount);
		this.setPageResults(list);
	}
	
	/**
	 * 总页数
	 * @param pageNo
	 */
	public void generate(int pageNo) {
		this.setPageNo(pageNo);
		//组装totalPage
		this.setTotalPage((int) Math.ceil((double) this.getTotalCount() / this.getPageSize()));
		if(this.getTotalPage() == 0) {
			this.setTotalPage(1);
		}
		//当前页传参溢出最大页，设置为当前页

		if(this.getPageNo() > this.getTotalPage()) {
			this.setPageNo(this.getTotalPage());
		}
		this.setStartRow((pageNo - 1) * this.getPageSize());
	}

	public List<T> getPageResults() {
		return pageResults;
	}

	public void setPageResults(List<T> pageResults) {
		this.pageResults = pageResults;
	}


	public Integer getPageNo() {
		return pageNo;
	}
	public void setPageNo(Integer pageNo) {
		this.pageNo = pageNo;
	}
	public Integer getTotalPage() {
		return totalPage;
	}

	public void setTotalPage(Integer totalPage) {
		this.totalPage = totalPage;
	}

	public Integer getTotalCount() {
		return totalCount;
	}

	public void setTotalCount(Integer totalCount) {
		this.totalCount = totalCount;
	}

	public void setTotalCountAndGenerate(Integer totalCount) {
		this.totalCount = totalCount;
		generate(pageNo);
	}
	
	public Integer getPageSize() {
		return pageSize;
	}

	public void setPageSize(Integer pageSize) {
		this.pageSize = pageSize;
	}

	public Integer getStartRow() {
		return startRow;
	}

	public void setStartRow(Integer startRow) {
		this.startRow = startRow;
	}



}
